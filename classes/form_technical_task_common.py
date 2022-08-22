from pandas import DataFrame

from classes.base_class import BaseClass
from configs import config, texts, cell_formats


class FormTechTaskComm(BaseClass):

    def __init__(self, pivot_table: DataFrame, row_numbers_list: list):
        super().__init__(pivot_table)
        self.factories_set = set()
        self.row_numbers_list = row_numbers_list  # номера строк для подсчёта итога по станции/сетям
        self.final_wb.filename = f'ТЗ_{self.budget_name}_{config.lot_name}_Общее.xlsx'
        self.final_ws.name = self.budget_name  # добавляет лист, в который будем записывать данные

    def big_table(self) -> None:
        """Получает все необходимые данные из сводной таблицы"""
        i = 1
        j = 0
        for row in self.temp_ws.iter_rows(min_row=2, max_row=self.temp_ws.max_row, min_col=2, max_col=19):
            self.table_data.append([i] + [cell.value for cell in row])
            if i < self.row_numbers_list[j]:
                i += 1
            else:
                i = 1
                j += 1
        [element.insert(5, None) for element in self.table_data]  # вставляет пустой столбец для технических требований

    def make_middle(self) -> None:
        """Добавляет данные в таблицу 1 ТЗ"""
        format_pivot_table = self.final_wb.add_format(cell_formats.format_pivot_table)
        quantity_format = self.final_wb.add_format(cell_formats.quantity_format)
        format_total_text = self.final_wb.add_format(cell_formats.format_total_text)
        format_total_num = self.final_wb.add_format(cell_formats.format_total_num)
        k = 0
        for row in self.table_data:
            factory_id = row[1]
            self.factories_set.add(factory_id)
            self.final_ws.write_number(f'A{self.row_number}', row[0], format_pivot_table)
            self.final_ws.write_row(f'B{self.row_number}', row[2:7], format_pivot_table)
            self.final_ws.write_formula(f'G{self.row_number}', f'=SUM(H{self.row_number}:T{self.row_number})',
                                        quantity_format)
            self.final_ws.write_row(f'H{self.row_number}', row[7:], quantity_format)
            self.final_ws.write_string(f'U{self.row_number}', texts.addresses[f'{factory_id}'], format_pivot_table)
            if row[0] == self.row_numbers_list[k]:
                self.row_number += 1
                self.final_ws.merge_range(f'A{self.row_number}:F{self.row_number}', texts.totals[f'{factory_id}'],
                                          format_total_text)
                for cell in config.cells:
                    self.final_ws.write_formula(f'{cell}{self.row_number}',
                                                f'=SUM({cell}{self.row_number - self.row_numbers_list[k]}:{cell}'
                                                f'{self.row_number - 1})',
                                                format_total_num)
                k += 1
                self.final_ws.write(f'U{self.row_number}', None, format_total_text)
            self.row_number += 1
        self.final_ws.merge_range(f'A{self.row_number}:F{self.row_number}', 'Общий итог', format_total_text)
        for cell in config.cells:
            next_total_row_number = 8
            formula_constructor = '='
            for index, total_row_number in enumerate(self.row_numbers_list):
                next_total_row_number += total_row_number
                formula_constructor += f'{cell}{next_total_row_number + index}+'
            self.final_ws.write_formula(f'{cell}{self.row_number}', formula_constructor[:-1], format_total_num)
        self.final_ws.write(f'U{self.row_number}', None, format_total_text)

    def make_tail(self) -> None:
        super().make_tail()
        addresses_string = ''
        for factory_id in self.factories_set:  # подготавливает список грузополучателей
            addresses_string += f'{texts.addresses[factory_id]}\n'
        self.final_ws.merge_range(f'E{self.row_number + 3}:U{self.row_number + 3}',
                                  texts.supply_conditions_desc2 + addresses_string.strip(), self.merge_format3)
        self.final_ws.set_row(self.row_number + 2, config.addresses_row_height[len(self.factories_set) - 1])

    def form(self) -> None:
        """Формирует ТЗ"""
        self.make_head()
        self.big_table()
        self.make_middle()
        self.make_tail()
        self.close_and_clear()
