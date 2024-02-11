import pathlib

from configs.config import sap_import_filename

ORIGIN_DATA_FILE_PATH = str(pathlib.Path().joinpath(f'../{sap_import_filename}'))
EXPECTED_DATA_FILE_PATH = str(pathlib.Path().joinpath('../expected_data'))
ACTUAL_DATA_FILE_PATH = str(pathlib.Path().joinpath('result'))
