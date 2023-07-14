If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "MC.8"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[6]").press
session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").selectedNode = "000017"
session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").doubleClickNode "000017"
session.findById("wnd[0]").sendVKey 12
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/ctxtSL_WERKS-LOW").text = "HC01"
session.findById("wnd[0]/usr/ctxtSL_SPMON-LOW").text = ""
session.findById("wnd[0]/usr/ctxtSL_SPMON-HIGH").text = ""
session.findById("wnd[0]/usr/ctxtSL_SPMON-HIGH").setFocus
session.findById("wnd[0]/usr/ctxtSL_SPMON-HIGH").caretPosition = 0
