import os

import pandas as pd
import pytest
from assertpy import assert_that

from tests.constants import ACTUAL_DATA_FILE_PATH, EXPECTED_DATA_FILE_PATH

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)


@pytest.mark.usefixtures('run_form_tech_task')
class TestCompareFiles:

    def setup_method(self):
        self.actual_file_names = os.listdir(ACTUAL_DATA_FILE_PATH)
        self.expected_file_names = os.listdir(EXPECTED_DATA_FILE_PATH)

    def test_compare_files_quantity(self):
        actual_quantity = len(self.actual_file_names)
        expected_quantity = len(self.expected_file_names)
        (assert_that(actual_quantity, description='Количество сгенерированных файлов не соответствует ожидаемому.')
         .is_equal_to(expected_quantity))

    def test_compare_file_names(self):
        for file_name in self.actual_file_names:
            (assert_that(self.expected_file_names, description=f'Имя файла "{file_name}" сгенерировалось с ошибкой.')
             .contains(file_name))

    def test_compare_files_data(self):
        for file_name in self.actual_file_names:
            actual_file_path = f'{ACTUAL_DATA_FILE_PATH}/{file_name}'
            expected_file_path = f'{EXPECTED_DATA_FILE_PATH}/{file_name}'
            actual_file = pd.read_excel(actual_file_path)
            expected_file = pd.read_excel(expected_file_path)

            diff = pd.concat([actual_file, expected_file]).drop_duplicates(keep=False).dropna()

            assert_that(diff, description=f'В файле "{file_name}" есть отличия.').is_empty()
