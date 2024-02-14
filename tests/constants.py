from pathlib import Path

from configs.config import sap_import_filename

ROOT = Path(__file__).parent
ORIGIN_DATA_FILE_PATH = str(ROOT.joinpath(sap_import_filename))
EXPECTED_DATA_FILE_PATH = str(ROOT.joinpath('expected_data'))
ACTUAL_DATA_FILE_PATH = str(ROOT.joinpath('result'))
