import warnings

from classes.form_engagement_report import FormEngagementReport
from classes.form_nmp_info import FormNmpInfo
from classes.form_technical_task_common import FormTechTaskComm
from classes.form_technical_task_sep import FormTechTaskSep
from configs import config
from utilities.helpers import engagement_report_helper, pivot_helper

warnings.simplefilter(action='ignore', category=FutureWarning)


def main(filepath: str = config.sap_import_filename):
    """Блок подготовки сводных таблиц."""
    common_tables = pivot_helper(filepath, 'common')
    sep_tables = pivot_helper(filepath, 'separated')
    nmp_info_tables = pivot_helper(filepath, 'nmp_info')
    engagement_report_table = engagement_report_helper(filepath)
    """Блок формирования ТЗ, сведений о НМЦ и отчёта по вовлечению"""
    for common_table in common_tables:
        FormTechTaskComm(common_table).form()  # формирует общие ТЗ
    for sep_table in sep_tables:  # формирует раздельные ТЗ
        FormTechTaskSep(sep_table).form()
    FormNmpInfo(nmp_info_tables).form()  # формирует сведения о НМЦ
    FormEngagementReport(engagement_report_table).form()  # формирует отчёт по вовлечению


if __name__ == "__main__":
    main()
