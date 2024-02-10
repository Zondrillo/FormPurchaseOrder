import warnings

from utilities.helpers import pivot_helper, engagement_report_helper
from configs import config
from classes.form_technical_task_common import FormTechTaskComm
from classes.form_technical_task_sep import FormTechTaskSep
from classes.form_nmp_info import FormNmpInfo
from classes.form_engagement_report import FormEngagementReport

warnings.simplefilter(action='ignore', category=FutureWarning)


def main(filepath: str = config.sap_import_filename):
    """Блок подготовки сводных таблиц"""
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
