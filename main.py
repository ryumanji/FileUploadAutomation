from smb.SMBConnection import SMBConnection
import win32com.client
import os, platform
import yaml
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support import expected_conditions
import pywinauto
from pywinauto.keyboard import SendKeys
import Log
from time import sleep
from webdriver_manager.chrome import ChromeDriverManager

class SMB():
    def __init__(self, info):
	self.filepath = ''
	self.directory = ''
        self.conn = self.connect_smb(info)

    def connect_smb(self, info):
        conn = SMBConnection(
            info['SMB']['UserName'],
            info['SMB']['Password'],
            platform.uname().node,
            info['SMB']['HOST'],)
        conn.connect(info['SMB']['IP'], 139)
        return conn

    def get_file_name(self):
        files = [f.filename for f in self.conn.listPath(self.directory, self.filepath)]
        return files

    def file_open(self, file):
        f = open(f'.\\self.filepath\\{file}', 'wb')
        return f

    def fetch_file(self, file):
        f = self.file_open(file)
        f.write(b'')
        self.conn.retrieveFile('***\\', f'**\\{file}', f)
        f.close()

    def remove_smb_file(self, file):
        f = self.file_open(file)
        self.conn.deleteFiles('***\\', f'**\\{file}', f)
        f.close()
        self.conn.close()

class Macro():
    """
    CSVファイルからエクセルファイル形式に変更する際、
    文字コードの変換が面倒だったためマクロ（csv2xlsx関数）で処理。
    Pythonにてエクセルファイルにパスワードを掛けるのが困難なため、
    こちらもマクロ（set_password関数）で処理。
    """
    def __init__(self):
        self.excel, self.wb = self.open_excel()

    def open_excel(self):
        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = True
        excel.DisplayAlerts = False
        filename = 'macro.xlsm'
        fullpath = os.path.join(os.getcwd(),'modules',filename)
        wb = excel.Workbooks.Open(Filename=fullpath)
        return excel, wb

    def execute_macro(self, file, new_file):
        pwd = os.getcwd()
        csv_path = os.path.join(f'{pwd}', 'files', f'{file}')
        xlsx_path = os.path.join(f'{pwd}', 'files', f'{new_file}')
        self.excel.Application.Run('csv2xlsx', csv_path, xlsx_path)
        self.excel.Application.Run('set_password', xlsx_path)

    def close_excel(self):
        self.wb.Close()
        self.excel.DisplayAlerts = True
        self.excel.Application.Quit()

class Web():
    """
    RedmineのナレッジベースはプラグインなのでAPIがない。
    そのため、Seleniumにてファイルをアップロードする。
    """
    def __init__(self, info):
        self.info = info
        self.driver = self.conn()

    def conn(self):
	path = ''
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')

        capabilities = DesiredCapabilities.CHROME.copy()
        capabilities['acceptInsecureCerts'] = True

        # 「この接続ではプライバシーが保護されません」対策
        dc = DesiredCapabilities.CHROME.copy()
        dc['acceptSslCerts'] = True

        driver = webdriver.Chrome(ChromeDriverManager().install(), desired_capabilities=dc)
        driver.get(self.info['Redmine']['UploadUrl'] + path)
        return driver

    def login(self):
        username = self.driver.find_element_by_id('username')
        username.send_keys(self.info['Redmine']['UserName'])
        password = self.driver.find_element_by_id('password')
        password.send_keys(self.info['Redmine']['Password'])
        login = self.driver.find_element_by_id('login-submit')
        login.click()

    def click_elements(self):
        upload = self.driver.find_element_by_class_name("add_attachment")
        upload.click()

    def select_file(self, upload_file):
        """
        アップロードするファイルの選択については、OS側の処理のため、
        pywinautoという外部モジュールを使用して処理。
        """
        findWindow = lambda: pywinauto.findwindows.find_windows(title=u'開く')[0]
        dialog = pywinauto.timings.wait_until_passes(5, 1, findWindow)
        pwa_app = pywinauto.Application()
        pwa_app.connect(handle=dialog)
        window = pwa_app[u"開く"]
        addres = window.children()[39]
        addres.click()
        dialog_dir = window.children()[43]

        tb = window[u"ファイル名(&N):"]
        pwd = os.getcwd()
        upload_file_path = os.path.join(f'{pwd}', 'files', f'{upload_file}')
        if tb.is_enabled():
            tb.click()
            edit = window.Edit4
            edit.set_focus()
            edit.type_keys(upload_file_path + '%O',with_spaces=True)

    def page_load(self):
        wait = WebDriverWait(self.driver, 10)
        elem = wait.until(expected_conditions.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[1]/div[3]/div[2]/form/input[4]')))
        elem.click()

    def close_window(self):
        self.driver.quit()

def main():
    # SMBよりCSVファイルを取得し、拡張子を .xlsx に変更
    smb.fetch_file(file)
    new_file = file.replace('.csv', '.xlsx')
    # マクロ関連処理
    macro.execute_macro(file, new_file)
    macro.close_excel()
    # WEB関連処理
    web.login()
    web.click_elements()
    web.select_file(new_file)
    sleep(5)
    web.page_load()
    web.close_window()
    # ファイルの削除
    smb.remove_smb_file(file)
    os.remove(os.path.join('files', file))
    os.remove(os.path.join('files', new_file))

def read_info():
    """
    Settings.yamlから各アカウント情報や、
    サーバーのホスト名、IPアドレス等の情報を引き出している。
    """
    with open('Settings.yaml', 'r') as f:
        info = yaml.load(f, Loader=yaml.SafeLoader)
    return info

if __name__ == "__main__":
    script_dir = os.path.dirname(__file__)
    os.chdir(script_dir)

    logger = Log.Log()
    logger.log_info('StartScript')

    info = read_info()
    smb = SMB(info)
    files = smb.get_file_name()

    if files == ['.', '..']:
        logger.log_info('CSVファイルがありません。')

    else:  # ディレクトリ内には '.', '..' の2つのファイルがある
        macro = Macro()
        web = Web(info)
        for file in files:
            if '.csv' not in file:
                continue
            else:
                main()
    logger.log_info('EndScript')