import xlsxwriter as xl
from pandas import DataFrame

from configs import config, texts, cell_formats
from utilities.helpers import get_supply_months


class BaseClass:

    def __init__(self, pivot_table: DataFrame):
        self.row_number = 8
        self.counter = 1  # счётчик для № позиции
        self.pivot_table = pivot_table
        self.pivot_table.fillna('', inplace=True)
        self.budget_name = pivot_table.index.get_level_values('Раздел_ГКПЗ')[0]  # получаем раздел ГКПЗ с которым работаем в данный момент
        self.current_factory = pivot_table.index.get_level_values('Завод')[0]  # получаем id завода, с которым работаем в данный момент
        self.final_wb = xl.Workbook()  # создаём конечный excel-файл, в который будем записывать данные
        self.final_ws = self.final_wb.add_worksheet()  # добавляем лист, в который будем записывать данные
        self.final_ws.set_landscape()  # альбомная ориентация
        self.final_ws.set_paper(9)  # формат А4
        self.final_ws.fit_to_pages(1, 0)  # вписать все столбцы на одну страницу
        self.final_ws.set_zoom(60)  # установить масштаб 60%
        self.final_ws.set_column('A:A', 6)
        self.final_ws.set_column('B:C', 13.5)
        self.final_ws.set_column('D:D', 43)
        self.final_ws.set_column('E:E', 54)
        self.final_ws.set_column('F:F', 9.5)
        self.final_ws.set_column('G:G', 18)
        self.final_ws.set_column('H:T', 15)
        self.final_ws.set_column('U:U', 46)

    def make_head(self) -> None:
        """Формирует шапку ТЗ"""
        format_head = self.final_wb.add_format(cell_formats.format_head)
        merge_format1 = self.final_wb.add_format(cell_formats.merge_format1)
        merge_format2 = self.final_wb.add_format(cell_formats.merge_format2)
        merge_format3 = self.final_wb.add_format(cell_formats.merge_format3)
        rotate = self.final_wb.add_format(cell_formats.rotate_format)
        self.final_ws.write('U1', 'Приложение № 2 к Приказу НФ "ПАО "Т Плюс"', format_head)
        self.final_ws.write('U2', '№___________________________________________ от ____________________________',
                            format_head)
        self.final_ws.merge_range('A4:U4', f'Техническое задание на поставку {config.lot_name}', merge_format2)
        self.final_ws.merge_range('A5:C5', 'Таблица 1', merge_format3)
        col_head = 0
        for element in config.tech_task_head:
            self.final_ws.merge_range(5, col_head, 6, col_head, element, merge_format1)
            col_head += 1
        months = get_supply_months()
        col_month = 7
        for month in months:
            self.final_ws.write_datetime(6, col_month, month, rotate)
            col_month += 1
        self.final_ws.merge_range('H6:T6', 'Срок поставки', merge_format1)
        self.final_ws.merge_range('U6:U7', 'Грузополучатель', merge_format1)

    def make_tail(self) -> None:
        """Вставляет таблицу № 2 в ТЗ"""
        self.row_number += 2
        merge_format1 = self.final_wb.add_format(cell_formats.merge_format3)
        merge_format2 = self.final_wb.add_format(cell_formats.merge_format1)
        self.merge_format3 = self.final_wb.add_format(cell_formats.merge_format4)
        self.final_ws.merge_range(f'A{self.row_number}:C{self.row_number}', 'Таблица 2', merge_format1)
        self.final_ws.write_string(f'A{self.row_number + 1}', '№ п/п', merge_format2)
        self.final_ws.merge_range(f'B{self.row_number + 1}:D{self.row_number + 1}', 'Показатель', merge_format2)
        self.final_ws.merge_range(f'E{self.row_number + 1}:U{self.row_number + 1}', 'Описание', merge_format2)
        self.final_ws.merge_range(f'A{self.row_number + 2}:A{self.row_number + 6}', 1, merge_format2)
        self.final_ws.merge_range(f'B{self.row_number + 2}:D{self.row_number + 6}', texts.supply_conditions_title,
                                  self.merge_format3)
        self.final_ws.merge_range(f'E{self.row_number + 2}:U{self.row_number + 2}', texts.supply_conditions_desc1,
                                  self.merge_format3)
        self.final_ws.set_row(self.row_number + 1, 60)
        self.final_ws.merge_range(f'E{self.row_number + 4}:U{self.row_number + 4}', texts.supply_conditions_desc3,
                                  self.merge_format3)
        self.final_ws.merge_range(f'E{self.row_number + 5}:U{self.row_number + 5}', texts.supply_conditions_desc4,
                                  self.merge_format3)
        self.final_ws.set_row(self.row_number + 4, 148.20)
        self.final_ws.merge_range(f'E{self.row_number + 6}:U{self.row_number + 6}', texts.supply_conditions_desc5,
                                  self.merge_format3)
        self.final_ws.merge_range(f'A{self.row_number + 7}:A{self.row_number + 10}', 2, merge_format2)
        self.final_ws.merge_range(f'B{self.row_number + 7}:D{self.row_number + 10}', texts.quality_requirements_title,
                                  self.merge_format3)
        self.final_ws.merge_range(f'E{self.row_number + 7}:U{self.row_number + 7}', texts.quality_requirements_desc1,
                                  self.merge_format3)
        self.final_ws.merge_range(f'E{self.row_number + 8}:U{self.row_number + 8}', texts.quality_requirements_desc2,
                                  self.merge_format3)
        self.final_ws.set_row(self.row_number + 7, 40)
        self.final_ws.merge_range(f'E{self.row_number + 9}:U{self.row_number + 9}', texts.quality_requirements_desc3,
                                  self.merge_format3)
        self.final_ws.merge_range(f'E{self.row_number + 10}:U{self.row_number + 10}', 'Иное: нет', self.merge_format3)
        self.final_ws.write_number(f'A{self.row_number + 11}', 3, merge_format2)
        self.final_ws.merge_range(f'B{self.row_number + 11}:D{self.row_number + 11}',
                                  texts.confirmation_of_compliance_title, self.merge_format3)
        self.final_ws.merge_range(f'E{self.row_number + 11}:U{self.row_number + 11}',
                                  texts.confirmation_of_compliance_desc1, self.merge_format3)
        self.final_ws.set_row(self.row_number + 10, 81)
        self.final_ws.merge_range(f'A{self.row_number + 12}:A{self.row_number + 14}', 4, merge_format2)
        self.final_ws.merge_range(f'B{self.row_number + 12}:D{self.row_number + 14}', texts.safety_requirements_title,
                                  self.merge_format3)
        self.final_ws.merge_range(f'E{self.row_number + 12}:U{self.row_number + 12}', texts.safety_requirements_desc1,
                                  self.merge_format3)
        self.final_ws.merge_range(f'E{self.row_number + 13}:U{self.row_number + 13}', texts.safety_requirements_desc2,
                                  self.merge_format3)
        self.final_ws.merge_range(f'E{self.row_number + 14}:U{self.row_number + 14}', 'Иное: нет', self.merge_format3)
        self.final_ws.merge_range(f'A{self.row_number + 15}:A{self.row_number + 18}', 5, merge_format2)
        self.final_ws.merge_range(f'B{self.row_number + 15}:C{self.row_number + 18}', 'Иные требования',
                                  self.merge_format3)
        self.final_ws.write_string(f'D{self.row_number + 15}', 'Эквивалент', self.merge_format3)
        self.final_ws.write_string(f'D{self.row_number + 16}', 'Толеранс (+/-), %', self.merge_format3)
        self.final_ws.write_string(f'D{self.row_number + 17}', 'Срок службы (расчетный ресурс)', self.merge_format3)
        self.final_ws.write_string(f'D{self.row_number + 18}', 'Другое', self.merge_format3)
        self.final_ws.merge_range(f'E{self.row_number + 15}:U{self.row_number + 15}', texts.equivalent_desc,
                                  self.merge_format3)
        self.final_ws.set_row(self.row_number + 14, 42.6)
        self.final_ws.merge_range(f'E{self.row_number + 16}:U{self.row_number + 16}', 'Нет', self.merge_format3)
        self.final_ws.merge_range(f'E{self.row_number + 17}:U{self.row_number + 17}', None, self.merge_format3)
        self.final_ws.merge_range(f'E{self.row_number + 18}:U{self.row_number + 18}', 'Нет', self.merge_format3)
