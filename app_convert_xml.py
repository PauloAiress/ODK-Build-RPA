# RPA routine for convert ODK xls form in ODK xml form
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import win32com.client
import sys

class ChromeInit:
    #setup chrome
    def __init__(self):
        #webdriver and this script should be in the same folder 
        self.driver_path = "./chromedriver.exe"
        self.options = webdriver.ChromeOptions()
        self.options.add_argument(
            "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36")
        self.options.add_experimental_option(
            "excludeSwitches", ["enable-automation"])
        self.options.add_experimental_option('useAutomationExtension', False)
        self.options.add_argument(
            '--disable-blink-features=AutomationControlled')

        self.chrome = webdriver.Chrome(
            self.driver_path,
            options=self.options
        )
        params = {'behavior': 'allow',
                  'downloadPath': r"G:\Drive\Teste\XML"} #replace this path with your own computer path
                                                         #this will be the standard download path while running this script
        self.chrome.execute_cdp_cmd('Page.setDownloadBehavior', params)

    def chose_file_btn(self):
        try:
            esc_arquivo = self.chrome.find_element(By.ID, 'id_file')
            esc_arquivo.send_keys(
                r'G:\Drive\ODK\Forms\Form_001.xlsx') #replace this path with your own computer path
                                                           #this will be the standard path where you'll save xlsforms
        except Exception as e:
            print(e)

    def submit_btn_click(self):
        try:
            submit = self.chrome.find_element(By.ID, 'submitBtn')
            submit.click()
        except Exception as e:
            print(e)

    def download_XMLForm(self):
        try:
            download = self.chrome.find_element(
                By.XPATH, '/html/body/div[1]/a[1]')
            download.click()
        except Exception as e:
            print(e)

    def acess(self, site):
        self.chrome.get(site)

    def exit(self):
        self.chrome.quit()

class Excel:
    def __init__(self):
        # This section will open Excel application
        self.xls = win32com.client.Dispatch("Excel.Application")

    def close_excel(self):
        #close excel
        self.xls.Quit()

class Sap:
    #setting up an opened SAP window, just for closing in the end of the script
    def __init__(self):
        try:            
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            if not type(SapGuiAuto) == win32com.client.CDispatch:
                return
            App = SapGuiAuto.GetScriptingEngine
            if not type(App) == win32com.client.CDispatch:
                SapGuiAuto = None
                return
            connection = App.Children(0)
            self.session = connection.Children(0)
            if not type(self.session) == win32com.client.CDispatch:
                connection = None
                App = None
                SapGuiAuto = None
                return
        
        except:
            print(sys.exc_info()[0])

    def exit(self):
        #This function closes SAP
        self.session.findById("wnd[0]/tbar[0]/okcd").Text = "/NEX"
        self.session.findById("wnd[0]").sendVKey(0)

if __name__ == '__main__':
    sap = Sap()
    excel = Excel()
    chrome = ChromeInit()
    
    chrome.acess("https://xlsform.getodk.org")
    sleep(2)
    chrome.chose_file_btn()
    sleep(2)
    chrome.submit_btn_click()
    sleep(2)
    chrome.download_XMLForm()
    sleep(3)
    chrome.exit()
    sleep(1)
    excel.close_excel()
    sleep(1)
    sap.exit()
