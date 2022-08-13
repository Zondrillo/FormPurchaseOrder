factories = ('7Q11', '7Q31', '7Q41', '7Q61', '7Q91', '7QB1')  # коды грузополучателей
budgets = ('РЕМОНТ', 'ЭКСПЛУАТАЦИЯ', 'ИП_ТПИР')  # перечень статей бюджета

kts_factories = ['7Q71', '7Q81', '7QA1']  # список заводов КТС, которые нужно унифицировать, т.е. заменить на 7Q61
dts_factories = ['7QC1', '7QF1']  # список заводов ДТС, которые нужно унифицировать, т.е. заменить на 7QB1
repair_budget = ['РЕМОНТ']  # бюджеты ремонта
exploitation_budget = ['ЭКСПЛУАТАЦИЯ АМОРТИЗ.ССП']  # бюджеты эксплуатации
investments_budget = ['ИП ТПИР', 'ИП ПИП']  # бюджеты инвестиций

year = 2023  # На какой год нужно сформировать ТЗ
start_month = 12  # С какого месяца начинаются поставки, 11 - Ноябрь, 12 - Декабрь и т.д.

crs = {
    'ЦРС ННовг Цех': '7Q11',
    'ЦРС Кстово Цехов': '7Q31',
    'ЦРС Дзержинск Цехов': '7Q41',
    'ЦРС Дзержинск РемРас': '7Q41',
    'НжФ ЦРС ТСКстово Цех': '7Q61',
    'НжФЦРСТСКстовРемРасх': '7Q61',
    'НжФ ЦРС ТСДзер Цех': '7QB1',
    'НжФЦРС ТСДзерРемРасх': '7QB1'
}
"""Наименования МВЗ для привязки ЦРС к станциям и сетям"""

columns = {
    'Раздел_ГКПЗ': '',
    'Завод': '',
    'Номер лота': '',
    '№ материала': '',
    'Краткий текст позиции': '',
    'Дата поставки': '',
    'ЕИ': '',
    'Количество': ''
}
"""Необходимые столбцы"""

head = ('№ п/п', '№ лота SAP', 'Код МТР SAP', 'Наименование продукции', 'Технические требования к продукции',
        'Ед. изм.', 'Количество ИТОГО')
"""Названия столбцов в шапке ТЗ"""

cells = 'GHIJKLMNOPQRST'
"""Ячейки для подсчёта итогов"""

lot_name = None
"""Переменная для хранения наименовая лота. Здесь она просто инициализируется, позже перезаписывается из хелпера."""

addresses_row_height = [46, 66, 86, 106, 126, 146]
"""Список для определения высоты строки с адресами грузополучатей в таблице №2"""

'----------------------------------------------------------------------------------------------------------------------'
"""Форматы ячеек"""

"""Общий формат, для наследования другими форматами"""
common_format = {
    'align': 'center',
    'valign': 'vcenter',
    'font': 'Tahoma',
    'font_size': 16,
    'border': True
}

"""Форматы для сводной таблицы"""
format_pivot_table = dict(common_format)
quantity_format = dict(common_format)
format_total_text = dict(common_format)
format_total_num = dict(common_format)

format_pivot_table.update(text_wrap=True)  # перенос слов
quantity_format.update(num_format='#,###0.000')
format_total_text.update(bold=True)  # полужирный шрифт
format_total_num.update(bold=True, num_format='#,###0.000')

"""Форматы для шапки ТЗ"""
format_head = dict(common_format)
merge_format1 = dict(common_format)
merge_format2 = dict(common_format)
merge_format3 = dict(common_format)
rotate_format = dict(common_format)

format_head.update(align='right', italic=True)  # выравнивание по правой стороне, курсивный шрифт
merge_format1.update(text_wrap=True)  # перенос слов
merge_format2.update(bold=True, border=False)  # полужирный шрифт, убирает границы ячеек
merge_format3.update(border=False)
rotate_format.update(rotation=90, num_format='mmmm yyyy')  # поворот содержимого ячейки на 90 градусов

"""Форматы для таблицы №2"""
merge_format4 = dict(common_format)

merge_format4.update(align='left', text_wrap=True)
