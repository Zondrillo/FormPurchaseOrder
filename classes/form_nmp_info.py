import xlsxwriter as xl

from configs import config, texts, cell_formats


class FormNmpInfo:

    def __init__(self, pivot_tables_list: list):
        self.row_number = 8
        self.counter = 1  # счётчик для № позиции
        self.current_factory = None
        self.current_budget = None
        self.factory_total_rows_numbers = []  # список с номерами строк итоговых значений по заводам
        self.budgets_total_rows_numbers = []  # список с номерами строк итоговых значений по бюджетам
        self.pivot_tables_list = pivot_tables_list
        self.final_wb = xl.Workbook(f'{config.results_dir_name}/{config.lot_name}/2_Сведения_о_НМЦ_'
                                    f'{config.lot_name}.xlsx')  # создаём конечный excel-файл
        self.final_ws = self.final_wb.add_worksheet()  # добавляем лист, в который будем записывать данные
        self.final_ws.set_portrait()  # книжная ориентация
        self.final_ws.set_paper(9)  # формат А4
        self.final_ws.fit_to_pages(1, 0)  # вписать все столбцы на одну страницу
        self.final_ws.set_zoom(100)  # установить масштаб 100%
        self.final_ws.set_column('A:A', 8.3)
        self.final_ws.set_column('B:B', 46.8)
        self.final_ws.set_column('C:C', 10.67)
        self.final_ws.set_column('D:D', 15.5)
        self.final_ws.set_column('E:E', 20.5)
        self.final_ws.set_column('F:F', 15.6)
        self.final_ws.set_column('G:G', 30)

    def make_head(self) -> None:
        """Формирует шапку НМЦ"""
        nmp_info_head_format = self.final_wb.add_format(cell_formats.nmp_info_head_format)
        nmp_info_columns_name_format = self.final_wb.add_format(cell_formats.nmp_info_columns_name_format)
        nmp_info_lot_name_format = self.final_wb.add_format(cell_formats.nmp_info_lot_name_format)
        nmp_info_title_format = self.final_wb.add_format(cell_formats.nmp_info_title_format)
        self.final_ws.write_string('G1', 'Приложение №2 к заявке на проведение закупки', nmp_info_head_format)
        self.final_ws.write_string('G2', '№___________________________________________ от ____________________________',
                                   nmp_info_head_format)
        self.final_ws.write_string('A4', texts.nmp_info_title, nmp_info_title_format)
        self.final_ws.write_string('A5', config.lot_name, nmp_info_lot_name_format)
        self.final_ws.write_row('A7', config.nmp_info_head, nmp_info_columns_name_format)

    def fill_table(self) -> None:
        """Формирует таблицу с данными"""
        num_format = self.final_wb.add_format(cell_formats.nmp_info_num_format)
        total_num_format = self.final_wb.add_format(cell_formats.nmp_info_total_num_format)
        position_number_format = self.final_wb.add_format(cell_formats.nmp_info_common_format)
        quantity_format = self.final_wb.add_format(cell_formats.nmp_info_quantity_format)
        for table in self.pivot_tables_list:  # Итерация по бюджетам
            self.factory_total_rows_numbers = []  # список с номерами строк итоговых значений по заводам
            self.current_budget = table.index.get_level_values('Раздел_ГКПЗ')[0]
            self.current_factory = table.index.get_level_values('Завод')[0]
            self.final_ws.merge_range(f'A{self.row_number}:G{self.row_number}',
                                      f'Бюджет "{self.current_budget}", {texts.factories_names[self.current_factory]}',
                                      total_num_format)
            self.row_number += 1
            self.counter = 1
            for row in table.itertuples():  # Итерация по заводам
                if (factory := row[0][1]) != self.current_factory:
                    """Если текущий завод не совпадает с предыдущим - тогда записывает итоги для этого завода и 
                    сбрасывает счётчик для № позиции"""
                    self.factory_total_rows_numbers.append(self.row_number)
                    self.write_factory_totals()
                    self.current_factory = factory
                    self.final_ws.merge_range(f'A{self.row_number + 1}:G{self.row_number + 1}',
                                              f'Бюджет "{self.current_budget}", '
                                              f'{texts.factories_names[self.current_factory]}',
                                              total_num_format)
                    self.row_number += 2
                    self.counter = 1
                price_with_vat = row[0][4] * (100 + config.vat_rate) / 100  # цена с НДС
                item_cost = price_with_vat * row[1]  # стоимость одной позиции с НДС
                line = list(row[0][2:]) + [price_with_vat]  # формирует строку с данными
                self.final_ws.write_number(f'A{self.row_number}', self.counter, position_number_format)  # № позиции
                self.final_ws.write_row(f'B{self.row_number}', line, num_format)  # наименование - начальная цена
                self.final_ws.write_number(f'F{self.row_number}', row[1], quantity_format)  # количество
                self.final_ws.write_number(f'G{self.row_number}', item_cost, num_format)  # стоимость позиции с НДС
                self.row_number += 1
                self.counter += 1
            self.factory_total_rows_numbers.append(self.row_number)
            self.budgets_total_rows_numbers.append(self.row_number + 1)
            """Записывает итоги по текущему бюджету"""
            self.write_factory_totals()
            self.row_number += 1
            self.write_budget_or_global_totals('subtotals')
            self.row_number += 1
        """Записывает общие итоги по всем бюджетам"""
        self.write_budget_or_global_totals('global_totals')

    def write_factory_totals(self) -> None:
        """Добавляет строки с итогами по каждой станции/сетям"""
        total_string_format = self.final_wb.add_format(cell_formats.nmp_info_total_string_format)
        total_num_format = self.final_wb.add_format(cell_formats.nmp_info_total_num_format)
        total_quantity_format = self.final_wb.add_format(cell_formats.nmp_info_total_quantity_format)
        self.final_ws.merge_range(f'A{self.row_number}:E{self.row_number}', texts.totals[self.current_factory],
                                  total_string_format)
        self.final_ws.write_formula(f'F{self.row_number}',
                                    f'=SUM(F{self.row_number - 1}:F{self.row_number - self.counter + 1})',
                                    total_quantity_format)
        self.final_ws.write_formula(f'G{self.row_number}',
                                    f'=SUM(G{self.row_number - 1}:G{self.row_number - self.counter + 1})',
                                    total_num_format)

    def write_budget_or_global_totals(self, total_type: str) -> None:
        """Записывает итоги по текущему бюджету или общие итоги по всем бюджетам"""
        if total_type == 'subtotals':
            string_format = self.final_wb.add_format(cell_formats.nmp_info_budget_string_total_format)
            total_format = self.final_wb.add_format(cell_formats.nmp_info_budget_total_format)
            quantity_format = self.final_wb.add_format(cell_formats.nmp_info_budget_quantity_total_format)
            total_rows_numbers = self.factory_total_rows_numbers
            data_string = f'Итого по бюджету "{self.current_budget}"'
        else:
            string_format = self.final_wb.add_format(cell_formats.nmp_info_global_string_total_format)
            total_format = self.final_wb.add_format(cell_formats.nmp_info_global_total_format)
            quantity_format = self.final_wb.add_format(cell_formats.nmp_info_global_quantity_total_format)
            total_rows_numbers = self.budgets_total_rows_numbers
            data_string = 'Итого по всем бюджетам'
        self.final_ws.merge_range(f'A{self.row_number}:E{self.row_number}', data_string, string_format)
        total_quantity_formula = "+".join([f'F{row_number}' for row_number in total_rows_numbers])
        total_cost_formula = "+".join([f'G{row_number}' for row_number in total_rows_numbers])
        self.final_ws.write_formula(f'F{self.row_number}', total_quantity_formula, quantity_format)
        self.final_ws.write_formula(f'G{self.row_number}', total_cost_formula, total_format)

    def form(self) -> None:
        """Формирует НМЦ"""
        self.make_head()
        self.fill_table()
        self.final_wb.close()
