import shutil

from pytest import fixture

from main import main
from tests.constants import ORIGIN_DATA_FILE_PATH


@fixture(scope='class')
def run_form_tech_task():
    yield main(filepath=ORIGIN_DATA_FILE_PATH)
    shutil.rmtree('result')
