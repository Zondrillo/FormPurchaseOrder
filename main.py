import warnings

from utilities.helpers import pivot_helper, engagement_report_helper
from configs import config
from classes.form_technical_task_common import FormTechTaskComm
from classes.form_technical_task_sep import FormTechTaskSep
from classes.form_nmp_info import FormNmpInfo
from classes.form_engagement_report import FormEngagementReport

warnings.simplefilter(action='ignore', category=FutureWarning)


if __name__ == "__main__":
    """Блок подготовки сводных таблиц"""
    common_tables = pivot_helper(config.sap_import_filename, 'common')
    sep_tables = pivot_helper(config.sap_import_filename, 'separated')
    nmp_info_tables = pivot_helper(config.sap_import_filename, 'nmp_info')
    engagement_report_table = engagement_report_helper(config.sap_import_filename)
    """Блок формирования ТЗ, сведений о НМЦ и отчёта по вовлечению"""
    for common_table in common_tables:
        FormTechTaskComm(common_table).form()  # формирует общие ТЗ
    for sep_table in sep_tables:  # формирует раздельные ТЗ
        FormTechTaskSep(sep_table).form()
    FormNmpInfo(nmp_info_tables).form()  # формирует сведения о НМЦ
    FormEngagementReport(engagement_report_table).form()  # формирует отчёт по вовлечению
