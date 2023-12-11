import sys, win32com.client, os, time, pandas as pd, datetime

with open(r'//mapdas04/GoPMO/Lumenis/Lumenis SAP/last_update.txt','r+') as file:
    file.truncate(0)

def write_log(message, file_path):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_message = f"[{timestamp}] {message}\n"
    with open(file_path, "a") as log_file:
        log_file.write(log_message)


def write_log2(message, file_path):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    last_update = f"{message} , [{timestamp}]\n"
    with open(file_path,"a") as log_file2:
        log_file2.write(last_update)



sid = '1. SAP ECC Production (PRD)'
  #1. SAP ECC Production (PRD)
  #3. SAP ECC Quality Assurance (QAS)
  #4. SAP ECC Business Simulation (BUS)

  

def SAPLogin1(): # overide consent prompt 
    try:
      
      os.system('TASKKILL /F /IM saplogon.exe')  
      os.startfile('C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe')
      time.sleep(1)
      SapGuiAuto = win32com.client.GetObject("SAPGUI")
      if not type(SapGuiAuto) == win32com.client.CDispatch:
        return
      application = SapGuiAuto.GetScriptingEngine
      if not type(application) == win32com.client.CDispatch:
        SapGuiAuto = None
        return
      connection = application.OpenConnection(sid,True)
      if not type(connection) == win32com.client.CDispatch:
        application = None
        SapGuiAuto = None
        return
 
    except:
      print(sys.exc_info())

    finally:
      connection = None
      application = None
      SapGuiAuto = None
    try:     
      SapGuiAuto = win32com.client.GetObject("SAPGUI")
      if not type(SapGuiAuto) == win32com.client.CDispatch:
        return
      application = SapGuiAuto.GetScriptingEngine
      if not type(application) == win32com.client.CDispatch:
        SapGuiAuto = None
        return
      connection = application.Children(0)
      if not type(connection) == win32com.client.CDispatch:
        application = None
        SapGuiAuto = None
        return
      session = connection.Children(0)
      if not type(session) == win32com.client.CDispatch:
        connection = None
        application = None
        SapGuiAuto = None
        return
     
      write_log(" - SAP data extraction started", r'//mapdas04/GoPMO/Lumenis/Lumenis SAP/run_log.txt')
      write_log(" - SAP data extraction started", r'//gsyklpstfile01.bsci.bossci.com/lumenis/TV_SCR/Tableau_Data/logs/run_log.txt')

      session.findById("wnd[0]").sendVKey (74)
      session.findById("wnd[0]/mbar/menu[4]/menu[12]").select()
      session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
      # os.system('TASKKILL /F /IM saplogon.exe')

      

    except:
      print(sys.exc_info())
      # print(sys.exc_info()[0])
      pass
    finally:
        # session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
        # session.findById("wnd[0]").sendVKey(0)
        session = None
        connection = None
        application = None
        SapGuiAuto = None

def SAPLogin2():
    try:
      
      # os.system('TASKKILL /F /IM saplogon.exe')  
      # os.startfile('C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe')
      time.sleep(1)
      
      SapGuiAuto = win32com.client.GetObject("SAPGUI")
      if not type(SapGuiAuto) == win32com.client.CDispatch:
        return

      application = SapGuiAuto.GetScriptingEngine
      if not type(application) == win32com.client.CDispatch:
        SapGuiAuto = None
        return

      connection = application.OpenConnection(sid,True)
      if not type(connection) == win32com.client.CDispatch:
        application = None
        SapGuiAuto = None
        return


    except:
      print(sys.exc_info()[0])

    finally:
      connection = None
      application = None
      SapGuiAuto = None

    try:     
      SapGuiAuto = win32com.client.GetObject("SAPGUI")
      if not type(SapGuiAuto) == win32com.client.CDispatch:
        return

      application = SapGuiAuto.GetScriptingEngine
      if not type(application) == win32com.client.CDispatch:
        SapGuiAuto = None
        return

      connection = application.Children(0)
      if not type(connection) == win32com.client.CDispatch:
        application = None
        SapGuiAuto = None
        return

      session = connection.Children(1)
      if not type(session) == win32com.client.CDispatch:
        connection = None
        application = None
        SapGuiAuto = None
        return

      write_log(" - Pass Cntrl+N", r'//mapdas04/GoPMO/Lumenis/Lumenis SAP/run_log.txt')
      write_log(" - Pass Cntrl+N", r'//gsyklpstfile01.bsci.bossci.com/lumenis/TV_SCR/Tableau_Data/logs/run_log.txt')
      # new block -start
      session.findById("wnd[0]/tbar[0]/okcd").text = "COOIS"
      session.findById("wnd[0]").sendVKey (0)
      session.findById("wnd[0]/tbar[1]/btn[17]").press()
      session.findById("wnd[1]/usr/txtV-LOW").text = "TECO_MH10"
      session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
      session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
      session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
      session.findById("wnd[1]").sendVKey (8)
      session.findById("wnd[0]").sendVKey (8)
      session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton ("&NAVIGATION_PROFILE_TOOLBAR_EXPAND")
      session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton ("&MB_EXPORT")
      session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem("&PC")
      session.findById("wnd[1]/tbar[0]/btn[0]").press()
      session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"//mapdas04/GoPMO/Lumenis/Lumenis SAP/"
      session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "COOIS-TECO.txt"
      session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 14
      session.findById("wnd[1]/tbar[0]/btn[11]").press()
      write_log2("COOIS", r'//mapdas04/GoPMO/Lumenis/Lumenis SAP/last_update.txt')
      # new block -end


      session.findById("wnd[0]/tbar[0]/okcd").text = "/NMDLD"
      session.findById("wnd[0]").sendVKey (0)
      session.findById("wnd[0]/tbar[1]/btn[17]").press()
      session.findById("wnd[1]/usr/txtV-LOW").text = "MH10_MRPMSG"
      session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
      session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
      session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
      session.findById("wnd[1]/tbar[0]/btn[8]").press()
      session.findById("wnd[0]/tbar[1]/btn[8]").press()
      session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").text = "MDLD"
      session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").setFocus()
      session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").caretPosition = 4
      session.findById("wnd[1]/tbar[0]/btn[13]").press()
      session.findById("wnd[0]/tbar[0]/okcd").text = "/nZM66"
      session.findById("wnd[0]").sendVKey (0)
      session.findById("wnd[0]/tbar[1]/btn[17]").press()
      session.findById("wnd[1]/usr/txtV-LOW").text = "MH10_AUTO"
      session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
      session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
      session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
      session.findById("wnd[1]/tbar[0]/btn[8]").press()
      session.findById("wnd[0]").sendVKey (8)
      session.findById("wnd[0]/mbar/menu[0]/menu[5]/menu[2]/menu[2]").select()
      session.findById("wnd[1]/tbar[0]/btn[0]").press()
      
      session.findById("wnd[1]/usr/ctxtDY_PATH").text = r'//mapdas04/GoPMO/Lumenis/Lumenis SAP/'
      

      session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "/zm66.txt"
      session.findById("wnd[1]/tbar[0]/btn[11]").press()
      session.findById("wnd[0]/tbar[0]/okcd").text = "/nSP01"
      session.findById("wnd[0]").sendVKey (0)
      write_log2("ZM66", r'//mapdas04/GoPMO/Lumenis/Lumenis SAP/last_update.txt')
     

      session.findById("wnd[0]/usr/tabsTABSTRIP_BL1/tabpSCR1/ssub%_SUBSCREEN_BL1:RSPOSP01NR:0100/txtS_RQTITL-LOW").text = "MDLD"
      session.findById("wnd[0]/usr/tabsTABSTRIP_BL1/tabpSCR1/ssub%_SUBSCREEN_BL1:RSPOSP01NR:0100/txtS_RQTITL-LOW").setFocus()
      session.findById("wnd[0]/usr/tabsTABSTRIP_BL1/tabpSCR1/ssub%_SUBSCREEN_BL1:RSPOSP01NR:0100/txtS_RQTITL-LOW").caretPosition = 4
      session.findById("wnd[0]/tbar[1]/btn[8]").press()
      session.findById("wnd[0]/usr/lbl[14,3]").setFocus()
      session.findById("wnd[0]/usr/lbl[14,3]").caretPosition = 0
      session.findById("wnd[0]").sendVKey (2)
      session.findById("wnd[0]/tbar[1]/btn[48]").press()
      session.findById("wnd[1]/tbar[0]/btn[0]").press()
      session.findById("wnd[1]/usr/ctxtDY_PATH").text = r'//mapdas04/GoPMO/Lumenis/Lumenis SAP/'
      session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MDLD.txt"
      session.findById("wnd[1]/tbar[0]/btn[11]").press()

      write_log2("MDLD", r'//mapdas04/GoPMO/Lumenis/Lumenis SAP/last_update.txt')
      write_log(" - SAP data extraction finished", r'//mapdas04/GoPMO/Lumenis/Lumenis SAP/run_log.txt')
      write_log(" - SAP data extraction finished", r'//gsyklpstfile01.bsci.bossci.com/lumenis/TV_SCR/Tableau_Data/logs/run_log.txt')

      

    except:
      # print(sys.exc_info())
      # print(sys.exc_info()[0])
      pass

    finally:
        # session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
        # session.findById("wnd[0]").sendVKey(0)
        # session = None
        connection = None
        application = None
        SapGuiAuto = None
        # write_log(" - SAP data extraction finished", r'//mapdas04/GoPMO/Lumenis/Lumenis SAP/run_log.txt')





SAPLogin1()
time.sleep(2)
SAPLogin2()
time.sleep(120)
os.system('TASKKILL /F /IM saplogon.exe') 