from pandas import DataFrame

from classes.base_class import BaseClass
from configs import config, texts


class FormTechTaskComm(BaseClass):

    def __init__(self, pivot_table: DataFrame, row_numbers_list: list):
        super().__init__(pivot_table)
        self.row_numbers_list = row_numbers_list  # номера строк для подсчёта итога по станции/сетям
        self.final_wb.filename = f'ТЗ_{self.budget_name}_{config.lot_name}_Общее.xlsx'
        self.final_ws.name = f'{self.budget_name}'  # добавляет лист, в который будем записывать данные

    def big_table(self) -> list:
        """Получает все необходимые данные из сводной таблицы"""
        lst = []
        i = 1
        j = 0
        for row in self.temp_ws.iter_rows(min_row=2, max_row=self.temp_ws.max_row, min_col=2, max_col=19):
            lst.append([i] + [cell.value for cell in row])
            if i < self.row_numbers_list[j]:
                i += 1
            else:
                i = 1
                j += 1
        [element.insert(5, None) for element in lst]  # вставляет пустой столбец для технических требований
        return lst

    def make_middle(self, lst_data: list) -> tuple:
        """Добавляет данные в таблицу 1 ТЗ"""
        r_num = 8
        format_pivot_table = self.final_wb.add_format(config.format_pivot_table)
        quantity_format = self.final_wb.add_format(config.quantity_format)
        format_total_text = self.final_wb.add_format(config.format_total_text)
        format_total_num = self.final_wb.add_format(config.format_total_num)
        self.final_ws.set_column('U:U', 46)
        k = 0
        factories_set = set()
        for row in lst_data:
            factory_id = row[1]
            factories_set.add(factory_id)
            self.final_ws.write_number(f'A{r_num}', row[0], format_pivot_table)
            self.final_ws.write_row(f'B{r_num}', row[2:7], format_pivot_table)
            self.final_ws.write_formula(r_num - 1, 6, f'=SUM(H{r_num}:T{r_num})', quantity_format)
            self.final_ws.write_row(f'H{r_num}', row[7:], quantity_format)
            self.final_ws.write_string(f'U{r_num}', texts.addresses[f'{factory_id}'], format_pivot_table)
            if row[0] == self.row_numbers_list[k]:
                self.final_ws.merge_range(r_num, 0, r_num, 5, texts.totals[f'{factory_id}'], format_total_text)
                for cell in config.cells:
                    self.final_ws.write_formula(f'{cell}{r_num + 1}',
                                                f'=SUM({cell}{r_num - self.row_numbers_list[k] + 1}:{cell}{r_num})',
                                                format_total_num)
                r_num += 1
                k += 1
                self.final_ws.write(f'U{r_num}', None, format_total_text)
            r_num += 1
        self.final_ws.merge_range(r_num - 1, 0, r_num - 1, 5, 'Общий итог', format_total_text)
        for cell in config.cells:
            next_total_row_number = 8
            formula_constructor = '='
            for index, total_row_number in enumerate(self.row_numbers_list):
                next_total_row_number += total_row_number
                formula_constructor += f'{cell}{next_total_row_number + index}+'
            self.final_ws.write_formula(f'{cell}{r_num}', formula_constructor[:-1], format_total_num)
        self.final_ws.write(f'U{r_num}', None, format_total_text)
        return r_num, factories_set

    def make_tail(self, row_num: int, factory_id_set: set) -> None:
        super().make_tail(row_num)
        addresses_string = ''
        for factory_id in factory_id_set:  # подготавливает список грузополучателей
            addresses_string += f'{texts.addresses[factory_id]}\n'
        self.final_ws.merge_range(f'E{row_num + 3}:U{row_num + 3}',
                                  texts.supply_conditions_desc2 + addresses_string.strip(), self.merge_format3)
        self.final_ws.set_row(row_num + 2, config.addresses_row_height[len(factory_id_set) - 1])

    def form(self) -> None:
        """Формирует ТЗ"""
        self.make_head()
        curr_row_num = self.make_middle(self.big_table())
        self.make_tail(curr_row_num[0] + 2, curr_row_num[1])
        self.close_and_clear()
