Attribute VB_Name = "Internet"
Public Sub vEmailMe(Message As String)
' Change the email addy to yours and the subject to whatever you need
Dim vShell As String
vShell = "Start.exe mailto:rompa@itson.nu?Subject="
vShell = vShell & Message
nResult = Shell(vShell, vbHide)
End Sub
