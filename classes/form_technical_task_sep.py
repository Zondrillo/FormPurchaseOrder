from pandas import DataFrame

from classes.base_class import BaseClass
from configs import config, texts


class FormTechTaskSep(BaseClass):

    def __init__(self, pivot_table: DataFrame):
        super().__init__(pivot_table)
        self.final_wb.filename = f'ТЗ_{self.factory_id}_{self.budget_name}_{config.lot_name}.xlsx'
        self.final_ws.name = f'{self.factory_id}'  # добавляет лист, в который будем записывать данные

    def big_table(self) -> list:
        """Получает все необходимые данные из сводной таблицы"""
        lst = []
        for index, row in enumerate(self.temp_ws.iter_rows(
                min_row=2, max_row=self.temp_ws.max_row, min_col=3, max_col=19), start=1):
            lst.append([index] + [cell.value for cell in row])
        [element.insert(4, None) for element in lst]  # вставляет пустой столбец для технических требований
        return lst

    def make_middle(self, lst: list, factory_id: str) -> int:
        """Добавляет данные в таблицу 1 ТЗ"""
        r_num = 8
        format_pivot_table = self.final_wb.add_format(config.format_pivot_table)
        quantity_format = self.final_wb.add_format(config.quantity_format)
        format_total_text = self.final_wb.add_format(config.format_total_text)
        format_total_num = self.final_wb.add_format(config.format_total_num)
        self.final_ws.set_column('U:U', 46)
        for row in lst:
            self.final_ws.write_row(f'A{r_num}', row[:6], format_pivot_table)
            self.final_ws.write_formula(r_num - 1, 6, f'=SUM(H{r_num}:T{r_num})', quantity_format)
            self.final_ws.write_row(f'H{r_num}', row[6:], quantity_format)
            self.final_ws.write_string(f'U{r_num}', texts.addresses[f'{factory_id}'], format_pivot_table)
            r_num += 1
        self.final_ws.merge_range(r_num - 1, 0, r_num - 1, 5, texts.totals[f'{factory_id}'], format_total_text)
        for cell in config.cells:
            self.final_ws.write_formula(f'{cell}{r_num}', f'=SUM({cell}8:{cell}{r_num - 1})', format_total_num)
        self.final_ws.write(f'U{r_num}', None, format_total_text)
        return r_num

    def make_tail(self, factory_id: str, row_num: int) -> int:
        super().make_tail(row_num)
        self.final_ws.merge_range(f'E{row_num + 3}:U{row_num + 3}',
                                  f'{texts.supply_conditions_desc2}{texts.addresses[factory_id]}', self.merge_format3)
        self.final_ws.set_row(row_num + 2, 46)
        return row_num + 20

    def signatory(self, factory_id, row_num):
        """Добавляет подписанта"""
        format1 = self.final_wb.add_format({'align': 'left', 'bold': True, 'font': 'Tahoma', 'font_size': 16})
        format1.set_align('bottom')
        self.final_ws.set_row(row_num - 1, 67.5)
        self.final_ws.write_string(f'A{row_num}', f'{texts.signatories[factory_id]}', format1)

    def form(self) -> None:
        """Формирует ТЗ"""
        self.make_head()
        curr_row_num = self.make_middle(self.big_table(), self.factory_id)
        row_for_sign = self.make_tail(self.factory_id, curr_row_num + 2)
        self.signatory(self.factory_id, row_for_sign)
        self.close_and_clear()
