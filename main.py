import warnings
import time

from utilities.helpers import pivot_helper
from configs import config
from classes.form_technical_task_common import FormTechTaskComm
from classes.form_technical_task_sep import FormTechTaskSep

warnings.simplefilter(action='ignore', category=FutureWarning)


if __name__ == "__main__":
    start = time.time()
    common_tables = pivot_helper('export.xlsx', 'common')
    sep_tables = pivot_helper('export.xlsx', 'separated')
    """Формируем список с количеством продукции для каждого бюджета и для каждого завода. Результатом является список 
    с вложенными списками следующего вида [Бюджет[Кол-во продукции по заводам]]"""
    quantity_budget = [[quantity for factory in config.factories
                        if (quantity := common_table.query(f'Завод == ["{factory}"]').shape[0]) != 0]
                       for common_table in common_tables]
    for i in range(len(common_tables)):  # формирует общие ТЗ
        FormTechTaskComm(common_tables[i], quantity_budget[i]).form()
    for sep_table in sep_tables:  # формирует раздельные ТЗ
        FormTechTaskSep(sep_table).form()
    print('Lead time: {:.2f} secs.'.format(time.time() - start))
