from pandas import DataFrame

from classes.base_class import BaseClass
from configs import config, texts, cell_formats


class FormTechTaskSep(BaseClass):

    def __init__(self, pivot_table: DataFrame):
        super().__init__(pivot_table)
        self.final_wb.filename = f'result/ТЗ_{self.current_factory}_{self.budget_name}_{config.lot_name}.xlsx'
        self.final_ws.name = self.current_factory  # добавляет лист, в который будем записывать данные

    def make_middle(self) -> None:
        """Добавляет данные в таблицу 1 ТЗ"""
        format_pivot_table = self.final_wb.add_format(cell_formats.format_pivot_table)
        quantity_format = self.final_wb.add_format(cell_formats.quantity_format)
        format_total_text = self.final_wb.add_format(cell_formats.format_total_text)
        format_total_num = self.final_wb.add_format(cell_formats.format_total_num)
        for row in self.pivot_table.itertuples():
            left_line = [self.counter] + list(row[0][2:5]) + [''] + [row[0][5]]  # подготавливает строку до G
            self.final_ws.write_row(f'A{self.row_number}', left_line, format_pivot_table)
            self.final_ws.write_formula(f'G{self.row_number}', f'=SUM(H{self.row_number}:T{self.row_number})',
                                        quantity_format)
            self.final_ws.write_row(f'H{self.row_number}', row[1:], quantity_format)
            self.final_ws.write_string(f'U{self.row_number}', texts.addresses[f'{self.current_factory}'],
                                       format_pivot_table)
            self.row_number += 1
            self.counter += 1
        self.final_ws.merge_range(f'A{self.row_number}:F{self.row_number}', texts.totals[f'{self.current_factory}'],
                                  format_total_text)
        for cell in config.cells:
            self.final_ws.write_formula(f'{cell}{self.row_number}', f'=SUM({cell}8:{cell}{self.row_number - 1})',
                                        format_total_num)
        self.final_ws.write_blank(f'U{self.row_number}', None, format_total_text)

    def make_tail(self) -> None:
        super().make_tail()
        self.final_ws.merge_range(f'E{self.row_number + 3}:U{self.row_number + 3}',
                                  f'{texts.supply_conditions_desc2}{texts.addresses[self.current_factory]}',
                                  self.merge_format3)
        self.final_ws.set_row(self.row_number + 2, 46)
        self.row_number += 20

    def signatory(self) -> None:
        """Добавляет подписанта"""
        signatory_format = self.final_wb.add_format(cell_formats.signatory_format)
        self.final_ws.set_row(self.row_number - 1, 67.5)
        self.final_ws.write_string(f'A{self.row_number}', f'{texts.signatories[self.current_factory]}',
                                   signatory_format)

    def form(self) -> None:
        """Формирует ТЗ"""
        self.make_head()
        self.make_middle()
        self.make_tail()
        self.signatory()
        self.final_wb.close()
