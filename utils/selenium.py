import tempfile
import shutil
import os
import random

from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

load_dotenv()

_driver = None
_temp_user_data_dir = None
_proxy_extension_temp_dir = None


def create_proxy_extension() -> str:
    random_str = ''.join(random.choices('abcdefghijklmnopqrstuvwxyz', k=8))
    
    PROXY_HOST = os.getenv('PROXY_HOST')
    PROXY_PORT = os.getenv('PROXY_PORT')
    PROXY_USER = os.getenv('PROXY_USER') + f'-session-{random_str}'
    PROXY_PASS = os.getenv('PROXY_PASS')
    
    manifest_json = '''
{
  "version": "1.0.0",
  "manifest_version": 2,
  "name": "Chrome Proxy",
  "permissions": [
    "proxy",
    "tabs",
    "unlimitedStorage",
    "storage",
    "<all_urls>",
    "webRequest",
    "webRequestBlocking"
  ],
  "background": {
    "scripts": ["background.js"]
  },
  "minimum_chrome_version": "22.0.0"
}
    '''
    background_js = f'''
var config = {{
    mode: "fixed_servers",
    rules: {{
      singleProxy: {{
        scheme: "http",
        host: "{PROXY_HOST}",
        port: parseInt({PROXY_PORT})
      }},
      bypassList: []
    }}
}};

chrome.proxy.settings.set({{value: config, scope: "regular"}}, function(){{}});

function callbackFn(details) {{
    return {{
        authCredentials: {{
            username: "{PROXY_USER}",
            password: "{PROXY_PASS}"
        }}
    }};
}}

chrome.webRequest.onAuthRequired.addListener(
    callbackFn,
    {{urls: ["<all_urls>"]}},
    ['blocking']
);
    '''
    temp_dir = tempfile.mkdtemp()
    extension_folder = os.path.join(temp_dir, "proxy_extension")
    os.makedirs(extension_folder)
    with open(os.path.join(extension_folder, "manifest.json"), "w") as manifest_file:
        manifest_file.write(manifest_json.strip())
    with open(os.path.join(extension_folder, "background.js"), "w") as bg_file:
        bg_file.write(background_js.strip())
    return extension_folder, temp_dir


def create_driver() -> webdriver.Chrome:
    global _driver, _temp_user_data_dir, _proxy_extension_temp_dir

    options = webdriver.ChromeOptions()

    _temp_user_data_dir = tempfile.mkdtemp()
    options.add_argument(f'--user-data-dir={_temp_user_data_dir}')

    if _proxy_extension_temp_dir:
        try:
            shutil.rmtree(_proxy_extension_temp_dir)
        except Exception as e:
            print(f'Error removing previous proxy extension temp dir: {e}')
        _proxy_extension_temp_dir = None

    extension_path, proxy_temp_dir = create_proxy_extension()
    _proxy_extension_temp_dir = proxy_temp_dir
    options.add_argument(f'--load-extension={extension_path}')

    _driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    _driver.maximize_window()

    return _driver


def quit_driver():
    global _driver, _temp_user_data_dir, _proxy_extension_temp_dir

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

    if _proxy_extension_temp_dir:
        try:
            shutil.rmtree(_proxy_extension_temp_dir)
        except Exception as e:
            print(f'Error removing proxy extension temp dir: {e}')
        _proxy_extension_temp_dir = None