# Import modules
import sys
import win32com.client
import psutil
import os




def sapconnection():
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

        # Check VC number
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "ZCP015"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = "BIOMASSAF"
        session.findById("wnd[1]/usr/txtENAME-LOW").text = "KFREIRE"
        session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
        session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 7
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/ctxtS_DT_ROM-HIGH").text = "31.12.2024"
        session.findById("wnd[0]/usr/ctxtS_DT_ROM-HIGH").setFocus()
        session.findById("wnd[0]/usr/ctxtS_DT_ROM-HIGH").caretPosition = 10
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/usr/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/shell").selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

    except:
        print(sys.exc_info()[0])
        print(sys.exc_info())

    finally:
        session = None
        connection = None
        application = None
        SapGuiAuto = None
