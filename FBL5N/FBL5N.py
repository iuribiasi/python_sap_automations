# -*- coding: utf-8 -*-
"""
Created on Tue Apr 19 11:14:45 2022

@author: I7770871
"""

# -*- coding: utf-8 -*-
"""
Created on Mon Apr 18 16:52:20 2022

@author: I7770871
"""

import win32com.client
import subprocess
import time
from datetime import datetime
import psutil
import pandas as pd




hoje = datetime.today().strftime('%d%m%Y')
hojenome = datetime.today().strftime('%Y.%m.%d')


class SapGui():
   
    def __init__(self):
        self.path = r'C:/Program Files (x86)/SAP/FrontEnd/SAPgui/saplogon.exe'
        self.sap_gui = subprocess.Popen(self.path)
        time.sleep(5)
        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not type(self.SapGuiAuto) == win32com.client.CDispatch:
            return
        application = self.SapGuiAuto.GetScriptingEngine
        self.connection = application.OpenConnection('01 - P42 - PCR/PPL (PRD)', True)
        self.session = self.connection.Children(0)
        time.sleep(3)
          
    def terminateSAP(self):
        self.sap_gui.kill()
        return 
 
                
    def FBL5NScript(self):
       
        self.session.findById("wnd[0]/tbar[0]/okcd").text="FBL5N"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/usr/ctxtDD_KUNNR-LOW").text="100000"
        self.session.findById("wnd[0]/usr/ctxtDD_KUNNR-HIGH").text="899999"
        self.session.findById("wnd[0]/usr/ctxtDD_KUNNR-HIGH").setFocus()
        self.session.findById("wnd[0]/usr/ctxtDD_KUNNR-HIGH").caretPosition=6
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").text="BC01"
        self.session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").setFocus()
        self.session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").caretPosition=4
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/usr/ctxtPA_STIDA").text=f"{hoje}"
        self.session.findById("wnd[0]/usr/ctxtPA_STIDA").setFocus()
        self.session.findById("wnd[0]/usr/ctxtPA_STIDA").caretPosition=8
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        self.session.findById("wnd[0]/usr/lbl[15,17]").setFocus()
        self.session.findById("wnd[0]/usr/lbl[15,17]").caretPosition=0
        self.session.findById("wnd[0]").sendVKey(16)
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text= r"C:\Users\I7770871\Desktop\AR_DIÁRIO"
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text=f"{hojenome}_AR.XLSX"
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition=16
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
               
                    
Sap_APP = SapGui()

Sap_APP.FBL5NScript()
Sap_APP.terminateSAP()


for proc in psutil.process_iter():
    if proc.name() == "EXCEL.EXE":
        proc.kill()


df = pd.read_excel(fr"C:\Users\I7770871\Desktop\AR_DIÁRIO\{hojenome}_AR.xlsx")
df.to_excel(f'//Qsbrprd/qliksense/BRASIL/OUTRAS FONTES/PCR_INDUSTRIAL/Finance/Daily_Contas_A_Receber/{hojenome}_AR.xlsx', index = False)
    