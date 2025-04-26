from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import tempfile
import shutil

_driver = None
_temp_user_data_dir = None


def create_driver() -> webdriver.Chrome:
    global _driver, _temp_user_data_dir

    options = webdriver.ChromeOptions()
    options.add_argument('--incognito')

    _temp_user_data_dir = tempfile.mkdtemp()
    options.add_argument(f'--user-data-dir={_temp_user_data_dir}')

    _driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    _driver.maximize_window()

    return _driver


def quit_driver():
    global _driver, _temp_user_data_dir

    if _driver:
        try:
            _driver.quit()
        except Exception as e:
            print(f'Error quitting driver: {e}')
        _driver = None

    if _temp_user_data_dir:
        try:
            shutil.rmtree(_temp_user_data_dir)
        except Exception as e:
            print(f'Error removing temp user data dir: {e}')
        _temp_user_data_dir = None