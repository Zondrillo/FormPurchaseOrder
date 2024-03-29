﻿from pandas import DataFrame

from classes.base_class import BaseClass
from configs import cell_formats, config, texts
from configs.config import factories


class FormTechTaskComm(BaseClass):

    def __init__(self, pivot_table: DataFrame):
        super().__init__(pivot_table)
        self.factories_set = set()
        self.factories_set.add(self.current_factory)
        self.factory_total_rows_numbers = []  # список с номерами строк итоговых значений по заводам
        self.final_wb.filename = f'result/ТЗ_{self.budget_name}_{config.lot_name}_Общее.xlsx'
        self.final_ws.name = self.budget_name  # добавляет лист, в который будет записывать данные

    def make_middle(self) -> None:
        """Добавляет данные в таблицу 1 ТЗ."""
        format_pivot_table = self.final_wb.add_format(cell_formats.format_pivot_table)
        quantity_format = self.final_wb.add_format(cell_formats.quantity_format)
        for row in self.pivot_table.itertuples():  # Итерация по заводам
            factory_code = row[0][1]
            if (factory := factories.get(factory_code)) != self.current_factory:
                """Если текущий завод не совпадает с предыдущим - тогда записывает итоги для этого завода и
                сбрасывает счётчик для № позиции"""
                self.factory_total_rows_numbers.append(self.row_number)
                self.write_totals('factory')
                self.current_factory = factory
                self.factories_set.add(factory)
                self.row_number += 1
                self.counter = 1
            left_line = [self.counter] + list(row[0][2:5]) + [''] + [row[0][5]]  # подготавливает строку до G
            self.final_ws.write_row(f'A{self.row_number}', left_line, format_pivot_table)
            self.final_ws.write_formula(f'G{self.row_number}', f'=SUM(H{self.row_number}:T{self.row_number})',
                                        quantity_format)
            self.final_ws.write_row(f'H{self.row_number}', row[1:], quantity_format)
            self.final_ws.write_string(f'U{self.row_number}', self.current_factory.address, format_pivot_table)
            self.row_number += 1
            self.counter += 1
        self.factory_total_rows_numbers.append(self.row_number)
        self.write_totals('factory')
        self.row_number += 1
        self.write_totals('global')

    def make_tail(self) -> None:
        super().make_tail()
        # подготавливает список адресов грузополучателей
        addresses = [factory.address for factory in sorted(self.factories_set, key=lambda x: x.code)]
        addresses_string = '\n'.join(addresses)
        self.final_ws.merge_range(f'E{self.row_number + 3}:U{self.row_number + 3}',
                                  texts.supply_conditions_desc2 + addresses_string, self.merge_format3)
        self.final_ws.set_row(self.row_number + 2, config.addresses_row_height[len(self.factories_set) - 1])

    def write_totals(self, total_type: str) -> None:
        """Добавляет строки с итогами по каждой станции/сетям."""
        format_total_text = self.final_wb.add_format(cell_formats.format_total_text)
        format_total_num = self.final_wb.add_format(cell_formats.format_total_num)
        if total_type == 'factory':
            self.final_ws.merge_range(f'A{self.row_number}:F{self.row_number}', self.current_factory.total,
                                      format_total_text)
            for cell in config.cells:
                self.final_ws.write_formula(
                    f'{cell}{self.row_number}',
                    f'SUM({cell}{self.row_number - 1}:{cell}{self.row_number - self.counter + 1})',
                    format_total_num)
        else:
            self.final_ws.merge_range(f'A{self.row_number}:F{self.row_number}', 'Общий итог', format_total_text)
            for cell in config.cells:
                total_quantity_formula = "+".join([f'{cell}{row_number}'
                                                   for row_number in self.factory_total_rows_numbers])
                self.final_ws.write_formula(f'{cell}{self.row_number}', total_quantity_formula, format_total_num)
        self.final_ws.write_blank(f'U{self.row_number}', None, format_total_text)

    def form(self) -> None:
        """Формирует ТЗ."""
        self.make_head()
        self.make_middle()
        self.make_tail()
        self.final_wb.close()
