Attribute VB_Name = "modEffects"
Option Explicit
'//Scrolling Credtis Const and Api Call
Public Const EM_GETLINECOUNT = &HBA
Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hwnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            lParam As Any) As Long   ' <---


'//Bouncing image variables
Public VelX As Integer
Public VelY As Integer
Public MaxX As Integer
Public MaxY As Integer
Public X As Integer
Public Y As Integer

'//Form On Top Code Starts Here
Public Declare Function SetWindowPos _
Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hwndinsertafter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long _
    ) As Long

'//SetWindowPos flags
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

'SetWindowsPos hwndInsertAfter values
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
   
Public mbOnTop As Boolean
'//Form OnTop Code Ends Here
