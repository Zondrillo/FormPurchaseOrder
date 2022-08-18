from pandas import DataFrame

from classes.base_class import BaseClass
from configs import config, texts


class FormTechTaskSep(BaseClass):

    def __init__(self, pivot_table: DataFrame):
        super().__init__(pivot_table)
        self.final_wb.filename = f'ТЗ_{self.factory_id}_{self.budget_name}_{config.lot_name}.xlsx'
        self.final_ws.name = self.factory_id  # добавляет лист, в который будем записывать данные

    def big_table(self) -> None:
        """Получает все необходимые данные из сводной таблицы"""
        for index, row in enumerate(self.temp_ws.iter_rows(
                min_row=2, max_row=self.temp_ws.max_row, min_col=3, max_col=19), start=1):
            self.table_data.append([index] + [cell.value for cell in row])
        [element.insert(4, None) for element in self.table_data]  # вставляет пустой столбец для технических требований

    def make_middle(self) -> None:
        """Добавляет данные в таблицу 1 ТЗ"""
        format_pivot_table = self.final_wb.add_format(config.format_pivot_table)
        quantity_format = self.final_wb.add_format(config.quantity_format)
        format_total_text = self.final_wb.add_format(config.format_total_text)
        format_total_num = self.final_wb.add_format(config.format_total_num)
        for row in self.table_data:
            self.final_ws.write_row(f'A{self.row_number}', row[:6], format_pivot_table)
            self.final_ws.write_formula(f'G{self.row_number}', f'=SUM(H{self.row_number}:T{self.row_number})',
                                        quantity_format)
            self.final_ws.write_row(f'H{self.row_number}', row[6:], quantity_format)
            self.final_ws.write_string(f'U{self.row_number}', texts.addresses[f'{self.factory_id}'], format_pivot_table)
            self.row_number += 1
        self.final_ws.merge_range(f'A{self.row_number}:F{self.row_number}', texts.totals[f'{self.factory_id}'],
                                  format_total_text)
        for cell in config.cells:
            self.final_ws.write_formula(f'{cell}{self.row_number}', f'=SUM({cell}8:{cell}{self.row_number - 1})',
                                        format_total_num)
        self.final_ws.write(f'U{self.row_number}', None, format_total_text)

    def make_tail(self) -> None:
        super().make_tail()
        self.final_ws.merge_range(f'E{self.row_number + 3}:U{self.row_number + 3}',
                                  f'{texts.supply_conditions_desc2}{texts.addresses[self.factory_id]}',
                                  self.merge_format3)
        self.final_ws.set_row(self.row_number + 2, 46)
        self.row_number += 20

    def signatory(self) -> None:
        """Добавляет подписанта"""
        format1 = self.final_wb.add_format({'align': 'left', 'bold': True, 'font': 'Tahoma', 'font_size': 16})
        format1.set_align('bottom')
        self.final_ws.set_row(self.row_number - 1, 67.5)
        self.final_ws.write_string(f'A{self.row_number}', f'{texts.signatories[self.factory_id]}', format1)

    def form(self) -> None:
        """Формирует ТЗ"""
        self.make_head()
        self.big_table()
        self.make_middle()
        self.make_tail()
        self.signatory()
        self.close_and_clear()
