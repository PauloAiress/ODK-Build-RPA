# configuração padrão para webscraping
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import win32com.client
import sys

class ChromeInit:
    def __init__(self, obra):
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
                  'downloadPath': r"G:\Meu Drive\Teste_ODK\XML" + obra}
        self.chrome.execute_cdp_cmd('Page.setDownloadBehavior', params)

    def botao_encontrar_arquivo(self, obra):
        divisao = obra[-3:]
        try:
            esc_arquivo = self.chrome.find_element(By.ID, 'id_file')
            esc_arquivo.send_keys(
                r'G:\Meu Drive\Teste_ODK\Forms' + obra + r'\Form-Apontamento-' + divisao + '.xlsx')
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
        # Essa função vai abrir o excel e executar a macro
        self.xls = win32com.client.Dispatch("Excel.Application")

    def close_excel(self):
        #fechar o excel
        self.xls.Quit()

class Sap:
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
        # Essa função vai fechar o SAP
        self.session.findById("wnd[0]/tbar[0]/okcd").Text = "/NEX"
        self.session.findById("wnd[0]").sendVKey(0)

if __name__ == '__main__':
    sap = Sap()
    excel = Excel()
    obras = [r'\329', r'\331', r'\339', r'\341', r'\342', r'\343', r'\345', r'\347']
    validador_obras = [False, True, False,
                       True, False,  False,  False,  False, ]
    for i in range(8):
        if validador_obras[i]:
            chrome = ChromeInit(obras[i])
            chrome.acess("https://xlsform.getodk.org")
            sleep(2)
            chrome.botao_encontrar_arquivo(obras[i])
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
