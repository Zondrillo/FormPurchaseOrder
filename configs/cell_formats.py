"""Форматы ячеек ТЗ."""

"""Общий формат, для наследования другими форматами"""
tech_task_common_format = {
    'align': 'center',
    'valign': 'vcenter',
    'font': 'Tahoma',
    'font_size': 16,
    'border': True
}

"""Форматы для сводной таблицы"""
format_pivot_table = dict(tech_task_common_format)
quantity_format = dict(tech_task_common_format)
format_total_text = dict(tech_task_common_format)
format_total_num = dict(tech_task_common_format)

format_pivot_table.update(text_wrap=True)  # перенос слов
quantity_format.update(num_format='#,###0.000')
format_total_text.update(bold=True)  # полужирный шрифт
format_total_num.update(bold=True, num_format='#,###0.000')

"""Форматы для шапки ТЗ"""
format_head = dict(tech_task_common_format)
merge_format1 = dict(tech_task_common_format)
merge_format2 = dict(tech_task_common_format)
merge_format3 = dict(tech_task_common_format)
rotate_format = dict(tech_task_common_format)

format_head.update(align='right', italic=True)  # выравнивание по правой стороне, курсивный шрифт
merge_format1.update(text_wrap=True)  # перенос слов
merge_format2.update(bold=True, border=False)  # полужирный шрифт, убирает границы ячеек
merge_format3.update(border=False)
rotate_format.update(rotation=90, num_format='mmmm yyyy')  # поворот содержимого ячейки на 90 градусов

"""Форматы для таблицы №2"""
merge_format4 = dict(tech_task_common_format)
signatory_format = dict(tech_task_common_format)

merge_format4.update(align='left', text_wrap=True)
signatory_format.update(align='left', valign='bottom', bold=True, border=False)

'----------------------------------------------------------------------------------------------------------------------'
"""Форматы ячеек НМЦ"""

"""Общий формат, для наследования другими форматами"""
nmp_info_common_format = {
    'align': 'center',
    'valign': 'vcenter',
    'font': 'Tahoma',
    'font_size': 10,
    'border': True
}

"""Форматы для шапки НМЦ"""
nmp_info_head_format = dict(nmp_info_common_format)
nmp_info_title_format = dict(nmp_info_common_format)
nmp_info_lot_name_format = dict(nmp_info_common_format)
nmp_info_columns_name_format = dict(nmp_info_common_format)

nmp_info_head_format.update(align='right', italic=True, border=False)
nmp_info_title_format.update(align='left', border=False)
nmp_info_lot_name_format.update(align='left', border=False, bold=True)
nmp_info_columns_name_format.update(text_wrap=True)

"""Форматы для таблицы с данными НМЦ"""
nmp_info_num_format = dict(nmp_info_common_format)
nmp_info_total_string_format = dict(nmp_info_common_format)
nmp_info_total_num_format = dict(nmp_info_common_format)
nmp_info_total_quantity_format = dict(nmp_info_common_format)
nmp_info_budget_string_total_format = dict(nmp_info_common_format)
nmp_info_budget_total_format = dict(nmp_info_common_format)
nmp_info_budget_quantity_total_format = dict(nmp_info_common_format)
nmp_info_global_string_total_format = dict(nmp_info_common_format)
nmp_info_global_total_format = dict(nmp_info_common_format)
nmp_info_global_quantity_total_format = dict(nmp_info_common_format)
nmp_info_quantity_format = dict(nmp_info_common_format)

nmp_info_num_format.update(num_format='#,##0.00')
nmp_info_total_string_format.update(align='left', bold=True)
nmp_info_total_num_format.update(num_format='#,##0.00', bold=True)
nmp_info_total_quantity_format.update(num_format='#,###0.000', bold=True)
nmp_info_budget_string_total_format.update(align='left', bold=True, font_size=12)
nmp_info_budget_total_format.update(num_format='#,##0.00', bold=True, font_size=12)
nmp_info_budget_quantity_total_format.update(num_format='#,###0.000', bold=True, font_size=12)
nmp_info_global_string_total_format.update(align='left', bold=True, font_size=14)
nmp_info_global_total_format.update(num_format='#,##0.00', bold=True, font_size=14)
nmp_info_global_quantity_total_format.update(num_format='#,###0.000', bold=True, font_size=14)
nmp_info_quantity_format.update(num_format='#,###0.000')

'----------------------------------------------------------------------------------------------------------------------'
"""Форматы ячеек отчёта по вовлечению"""

"""Общий формат, для наследования другими форматами"""
engagement_report_common_format = {
    'align': 'center',
    'valign': 'vcenter',
    'font': 'Tahoma',
    'font_size': 10,
    'border': True
}

"""Форматы для шапки отчёта по вовлечению"""
engagement_report_head_format1 = dict(engagement_report_common_format)
engagement_report_head_format2 = dict(engagement_report_common_format)
engagement_report_title_format = dict(engagement_report_common_format)
engagement_report_lot_name_format = dict(engagement_report_common_format)
engagement_report_lot_num_format = dict(engagement_report_common_format)
engagement_report_columns_name_format = dict(engagement_report_common_format)

engagement_report_head_format1.update(align='right', border=False)
engagement_report_head_format2.update(border=False, text_wrap=True)
engagement_report_title_format.update(border=False, bold=True, font_size=12)
engagement_report_lot_name_format.update(align='left', border=False, bold=True)
engagement_report_lot_num_format.update(border=False)
engagement_report_columns_name_format.update(text_wrap=True)

"""Форматы для таблицы с данными отчёта по вовлечению"""
engagement_report_price_format = dict(engagement_report_common_format)
engagement_report_quantity_format = dict(engagement_report_common_format)
engagement_report_total_string_format = dict(engagement_report_common_format)
engagement_report_bottom_border_format = dict(engagement_report_common_format)

engagement_report_price_format.update(num_format='#,##0.00')
engagement_report_quantity_format.update(num_format='#,###0.000')
engagement_report_total_string_format.update(align='left', bold=True)
engagement_report_bottom_border_format.update(border=False, bottom=True)
