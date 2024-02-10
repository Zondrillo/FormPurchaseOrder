import os

from assertpy import assert_that

from tests.constants import ACTUAL_DATA_FILE_PATH, EXPECTED_DATA_FILE_PATH


class TestCompareFiles:

    def test_compare_files_quantity(self, run_form_tech_task):
        actual_quantity = len(os.listdir(ACTUAL_DATA_FILE_PATH))
        expected_quantity = len(os.listdir(EXPECTED_DATA_FILE_PATH))
        (assert_that(actual_quantity, description='Количество сгенерированных файлов не соответсвует ожидаемому.')
         .is_equal_to(expected_quantity))
