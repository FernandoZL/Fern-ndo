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
session.findById("wnd[0]/usr/lbl[82,17]").setFocus
session.findById("wnd[0]/usr/lbl[82,17]").caretPosition = 3
session.findById("wnd[0]").sendVKey 12
session.findById("wnd[0]").sendVKey 12
session.findById("wnd[0]").sendVKey 12
session.findById("wnd[0]").sendVKey 12
