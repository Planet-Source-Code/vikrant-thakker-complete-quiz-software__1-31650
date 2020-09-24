Attribute VB_Name = "EnableDisableCtrlAltDel"
'Disabling Ctrl - Alt - Delete And Ctrl - Esc
'Declarations
Private Declare Function SystemParametersInfo Lib _
"user32" Alias "SystemParametersInfoA" (ByVal uAction _
As Long, ByVal uParam As Long, ByVal lpvParam As Any, _
ByVal fuWinIni As Long) As Long
'code For Diasabling Ctrl-Alt-Del and Ctrl-Esc
Public Sub DisableCtrlAltDelete(bDisabled As Boolean)
    Dim X As Long
    X = SystemParametersInfo(97, bDisabled, CStr(1), 0)
End Sub



