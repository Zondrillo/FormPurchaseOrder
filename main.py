import warnings
import os
from pathlib import Path

from shutil import make_archive, rmtree
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename

from utilities.helpers import pivot_helper, engagement_report_helper
from configs import config
from classes.form_technical_task_common import FormTechTaskComm
from classes.form_technical_task_sep import FormTechTaskSep
from classes.form_nmp_info import FormNmpInfo
from classes.form_engagement_report import FormEngagementReport

warnings.simplefilter(action='ignore', category=FutureWarning)

app = Flask(__name__)
app.secret_key = "caircocoders-ednalan-2020"
app.config['UPLOAD_FOLDER'] = 'uploads'


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        for file in request.files.getlist("file"):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        for root, dirs, files in os.walk(app.config['UPLOAD_FOLDER']):
            for file in files:
                file_path = os.path.join(root, file)
                main(file_path)
                Path(file_path).unlink(missing_ok=True)
        make_archive(config.results_dir_name, 'zip', config.results_dir_name)
        rmtree(os.path.join(config.results_dir_name), ignore_errors=True)
    return render_template('index.html')


@app.route('/download')
def download_file():
    path = 'output.zip'
    return send_file(path, as_attachment=True)


@app.route('/check')
def check():
    return str(Path('output.zip').is_file())


def main(file_name):
    """Блок подготовки сводных таблиц"""
    common_tables = pivot_helper(file_name, 'common')
    sep_tables = pivot_helper(file_name, 'separated')
    nmp_info_tables = pivot_helper(file_name, 'nmp_info')
    engagement_report_table = engagement_report_helper(file_name)
    Path(f'{config.results_dir_name}/{config.lot_name}').mkdir(parents=True, exist_ok=True)
    """Блок формирования ТЗ, сведений о НМЦ и отчёта по вовлечению"""
    for common_table in common_tables:
        FormTechTaskComm(common_table).form()  # формирует общие ТЗ
    for sep_table in sep_tables:  # формирует раздельные ТЗ
        FormTechTaskSep(sep_table).form()
    FormNmpInfo(nmp_info_tables).form()  # формирует сведения о НМЦ
    FormEngagementReport(engagement_report_table).form()  # формирует отчёт по вовлечению


if __name__ == "__main__":
    app.run(debug=True)
