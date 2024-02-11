import xlsxwriter as xl
from pandas import DataFrame

from configs import cell_formats, config


class FormEngagementReport:

    def __init__(self, df: DataFrame):
        self.data = df
        self.row_number = 11
        # создаём конечный excel-файл
        self.final_wb = xl.Workbook(f'result/3_Отчёт_по_вовлечению_{config.lot_name}.xlsx')
        self.final_ws = self.final_wb.add_worksheet()  # добавляем лист, в который будем записывать данные
        self.final_ws.set_landscape()  # альбомная ориентация
        self.final_ws.set_paper(9)  # формат А4
        self.final_ws.fit_to_pages(1, 0)  # вписать все столбцы на одну страницу
        self.final_ws.set_zoom(85)  # установить масштаб 85%
        self.final_ws.set_column('A:A', 10)
        self.final_ws.set_column('B:C', 8)
        self.final_ws.set_column('D:F', 11)
        self.final_ws.set_column('G:G', 12.5)
        self.final_ws.set_column('H:H', 43.5)
        self.final_ws.set_column('I:I', 9.5)
        self.final_ws.set_column('J:J', 13.8)
        self.final_ws.set_column('K:K', 11)
        self.final_ws.set_column('L:L', 14)
        self.final_ws.set_column('M:M', 11)
        self.final_ws.set_column('N:N', 14)
        self.final_ws.set_column('O:O', 11)
        self.final_ws.set_column('P:P', 14)
        self.final_ws.set_row(1, 28)
        self.final_ws.set_row(8, 95)
        self.final_ws.freeze_panes(10, 0)

    def make_head(self) -> None:
        """Формирует шапку отчёта по вовлечению."""
        head_format1 = self.final_wb.add_format(cell_formats.engagement_report_head_format1)
        head_format2 = self.final_wb.add_format(cell_formats.engagement_report_head_format2)
        title_format = self.final_wb.add_format(cell_formats.engagement_report_title_format)
        lot_name_format = self.final_wb.add_format(cell_formats.engagement_report_lot_name_format)
        lot_num_format = self.final_wb.add_format(cell_formats.engagement_report_lot_num_format)
        columns_name_format = self.final_wb.add_format(cell_formats.engagement_report_columns_name_format)
        self.final_ws.merge_range('L1:P1', 'Приложение №2 к Положению', head_format1)
        self.final_ws.merge_range('M2:P2',
                                  'Утверждено _____________________(ФИО, должность инициатора закупки на ЗК/ЦЗО)',
                                  head_format2)
        self.final_ws.merge_range('A3:P3', 'Отчет по вовлечению МТР к заявке на ЦЗО/закупочной комиссии филиала',
                                  title_format)
        self.final_ws.merge_range('A5:G5', 'Филиал "Нижегородский"', lot_name_format)
        self.final_ws.merge_range('A6:G6', f'Закупка {config.lot_name}', lot_name_format)
        self.final_ws.write_string('C7', 'номер лота ГКПЗ и наименование закупки', lot_num_format)
        for cell in 'ABCDEFGHIJ':
            self.final_ws.merge_range(f'{cell}8:{cell}9', None, columns_name_format)
        self.final_ws.write_row('A8', config.engagement_report_head_left, columns_name_format)
        for string, cells in config.engagement_report_head_right_top.items():
            self.final_ws.merge_range(cells, string, columns_name_format)
        self.final_ws.write_row('K9', config.engagement_report_head_right_bottom * 3, columns_name_format)
        self.final_ws.write_row('A10', range(1, 17), columns_name_format)

    def fill_table(self) -> None:
        """Формирует таблицу с данными."""
        table_format = self.final_wb.add_format(cell_formats.engagement_report_common_format)
        price_format = self.final_wb.add_format(cell_formats.engagement_report_price_format)
        quantity_format = self.final_wb.add_format(cell_formats.engagement_report_quantity_format)
        total_string_format = self.final_wb.add_format(cell_formats.engagement_report_total_string_format)
        bottom_border_format = self.final_wb.add_format(cell_formats.engagement_report_bottom_border_format)
        simple_format = self.final_wb.add_format({'font': 'Tahoma', 'font_size': 10})
        for row in self.data.itertuples():
            price = row[9]
            quantity = row[10]
            cost = price * quantity
            self.final_ws.write_row(f'A{self.row_number}', list(row[:9]), table_format)
            self.final_ws.write_number(f'J{self.row_number}', price, price_format)
            self.final_ws.write_number(f'K{self.row_number}', quantity, quantity_format)
            self.final_ws.write_number(f'L{self.row_number}', cost, price_format)
            self.final_ws.write_row(f'M{self.row_number}:N{self.row_number}', ' ' * 2, table_format)
            self.final_ws.write_number(f'O{self.row_number}', quantity, quantity_format)
            self.final_ws.write_number(f'P{self.row_number}', cost, price_format)
            self.row_number += 1
        self.final_ws.merge_range(f'A{self.row_number}:J{self.row_number}', 'ИТОГО', total_string_format)
        for index, cell in enumerate('KLMNOP'):
            cell_format = quantity_format if (index % 2) == 0 else price_format
            self.final_ws.write_formula(
                f'{cell}{self.row_number}',
                f'SUM({cell}{self.row_number - 1}:{cell}{self.row_number - self.data.shape[0]})',
                cell_format)
        self.row_number += 2
        self.final_ws.write_row(f'A{self.row_number}', ' ' * 8, bottom_border_format)
        self.final_ws.write_string(f'J{self.row_number}', 'Дата проведения вовлечения', simple_format)
        self.final_ws.write_string(f'L{self.row_number}', '', bottom_border_format)
        self.row_number += 1
        self.final_ws.write_string(f'A{self.row_number}', 'Заместитель директора филиала по логистике и закупкам',
                                   simple_format)

    def form(self) -> None:
        """Формирует отчёт по вовлечению."""
        self.make_head()
        self.fill_table()
        self.final_wb.close()
