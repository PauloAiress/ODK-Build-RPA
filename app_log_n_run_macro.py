import win32com.client
import sys
import os, os.path
import subprocess
from time import sleep

#rotina de RPA para abrir o SAP e retirar relatórios de equipamentos estoque

class Sap:
    def __init__(self):
    # Essa função vai abrir o SAP
        try:
            path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
            subprocess.Popen(path)
            sleep(3)
            
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            if not type(SapGuiAuto) == win32com.client.CDispatch:
                return
            
            App = SapGuiAuto.GetScriptingEngine
            if not type(App) == win32com.client.CDispatch:
                SapGuiAuto = None
                return
            # trocar DESCRIÇÃO pelo nome da conexão
            connection = App.OpenConnection("PRD - Produtivo", True)
            
            if not type(connection) == win32com.client.CDispatch:
                App = None
                SapGuiAuto = None
                return

            self.session = connection.Children(0)
            if not type(self.session) == win32com.client.CDispatch:
                connection = None
                App = None
                SapGuiAuto = None
                return
        
        except:
            print(sys.exc_info()[0])
            
    def saplogin(self):
        # essa função vai logar no SAP
        # trocar usuário e senha pelo seu usuário e senha
        self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "pauloos"
        self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "una234"
        self.session.findById("wnd[0]").sendVKey(0)

class Excel:
    def __init__(self):
        # Essa função vai abrir o excel e executar a macro
        self.xls = win32com.client.Dispatch("Excel.Application")

    def run_macro(self):
        # Abri o arquivo conforme o caminho passado
        self.xls.Workbooks.Open(os.path.abspath(r"G:\Meu Drive\Teste_ODK\vba_py\Macros_ODK.xlsm"), ReadOnly = 1)
        self.xls.Application.Run("Macros_ODK.xlsm!Módulo3.IE03_MB52")
    
    def close_excel(self):
        #fechar o excel
        self.xls.Quit()

if __name__ == '__main__':
    sap = Sap()
    excel = Excel()
    sap.saplogin()
    sleep(2)
    excel.run_macro()
    sleep(2)


