import warnings

from utilities.helpers import pivot_helper, engagement_report_helper
from configs import config
from classes.form_technical_task_common import FormTechTaskComm
from classes.form_technical_task_sep import FormTechTaskSep
from classes.form_nmp_info import FormNmpInfo
from classes.form_engagement_report import FormEngagementReport

warnings.simplefilter(action='ignore', category=FutureWarning)


if __name__ == "__main__":
    common_tables = pivot_helper(config.sap_import_filename, 'common')
    sep_tables = pivot_helper(config.sap_import_filename, 'separated')
    nmp_info_tables = pivot_helper(config.sap_import_filename, 'nmp_info')
    engagement_report_table = engagement_report_helper(config.sap_import_filename)
    """Формируем список с количеством продукции для каждого бюджета и для каждого завода. Результатом является список
    с вложенными списками следующего вида [Бюджет[Кол-во продукции по заводам]]"""
    quantity_budget = [[quantity for factory in config.factories
                        if (quantity := common_table.query(f'Завод == ["{factory}"]').shape[0]) != 0]
                       for common_table in common_tables]
    for i in range(len(common_tables)):  # формирует общие ТЗ
        FormTechTaskComm(common_tables[i], quantity_budget[i]).form()
    for sep_table in sep_tables:  # формирует раздельные ТЗ
        FormTechTaskSep(sep_table).form()
    FormNmpInfo(nmp_info_tables).form()  # формирует сведения о НМЦ
    FormEngagementReport(engagement_report_table).form()  # формирует отчёт по вовлечению
