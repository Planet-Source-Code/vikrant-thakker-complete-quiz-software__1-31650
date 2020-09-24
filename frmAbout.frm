VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Q"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   60
      TabIndex        =   3
      Top             =   2940
      Width           =   3945
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK...   I'm Bored, Can I Go Now"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   180
         Left            =   60
         TabIndex        =   4
         Top             =   30
         Width           =   3840
      End
   End
   Begin VB.Timer tmrFlyIn 
      Interval        =   500
      Left            =   3330
      Top             =   435
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000002&
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   -75
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   4155
   End
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   3345
      Top             =   -30
   End
   Begin VB.PictureBox picScroll 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   2505
      Left            =   30
      ScaleHeight     =   2505
      ScaleWidth      =   3975
      TabIndex        =   1
      Top             =   210
      Width           =   3975
      Begin VB.TextBox txtScroll 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   360
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2520
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private s As Integer

'//++++++++++++++++ Form Load Code ++++++++++++++++
Private Sub Form_Load()

Call HideTaskBar
Call HideDesktop

'Variable Declarations
    Dim iFileNum As Integer
    Dim lLineCount As Long
    Dim lLineHeight As Long
    Call HideDesktop
    On Error GoTo ErrHandler 'Goto to ErrHandler if an error occurs
    
        iFileNum = FreeFile
        'open file and read text from it
        Open App.Path & "\Credits.txt" For Input As iFileNum
        txtScroll = Input(LOF(iFileNum), iFileNum)
        Close #iFileNum 'close file
        'lLineCount = SendMessage(txtScroll.hwnd, EM_GETLINECOUNT, 0&, 0&)
        'lLineHeight = TextHeight("Test") 'Get the height of text in file
        txtScroll.Height = 5000 'lLineHeight * lLineCount + 700
        picScroll.Left = 0
        picScroll.Visible = True
        tmrScroll.Enabled = True
        OnTop = True

Dim ctl As Control
'//This disables all controls in frmMain to stop user
'//clicking anything when this form is actve
For Each ctl In frmMain.Controls
    If (TypeOf ctl Is Frame) Or (TypeOf ctl Is CommandButton) _
    Or (TypeOf ctl Is TextBox) Then
        ctl.Enabled = False
    End If
Next

Exit Sub
ErrHandler:
    txtScroll.Text = "File Not Found !!!" & vbNewLine & "The Credits.txt file is missing"
    Resume Next
End Sub
'//+++++++++++++++++ End Form Load Code +++++++++++

'//+++++++++++++++ Unload Code ++++++++++++++++++++
Private Sub cmdOK_Click()
Dim ctl As Control
'//enables all controls on frmMain again
For Each ctl In frmMain.Controls
    If (TypeOf ctl Is Frame) Or (TypeOf ctl Is CommandButton) _
    Or (TypeOf ctl Is TextBox) Then
        ctl.Enabled = True
    End If
Next
    Unload Me 'unload form
End Sub

Private Sub Form_Terminate()
Call ShowDesktop
Call ShowTaskBar
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call ShowDesktop
Call ShowTaskBar
End Sub

Private Sub Frame1_Click()
'//enables all controls on frmMain again
Dim ctl As Control
For Each ctl In frmMain.Controls
    If (TypeOf ctl Is Frame) Or (TypeOf ctl Is CommandButton) _
    Or (TypeOf ctl Is TextBox) Then
        ctl.Enabled = True
    End If
Next
    Unload Me
End Sub

Private Sub Label1_Click()
'//enables all controls on frmMain again
Dim ctl As Control
For Each ctl In frmMain.Controls
    If (TypeOf ctl Is Frame) Or (TypeOf ctl Is CommandButton) _
    Or (TypeOf ctl Is TextBox) Then
        ctl.Enabled = True
    End If
Next
    frmAbout.Hide
    'frmInfo.Show
    frmMain.Show
    'Unload Me
End Sub
'//++++++++++++ End Unload Code ++++++++++++++++++++


'//+++++++++++++ tmrFlyIn Code ++++++++++++++++++++
Private Sub tmrFlyIn_Timer()
    s = s + 1
    Me.Caption = Mid("Quiz 2001 Ver.3.0 ", 1, s)
    
    If s = 40 Then
        Me.Caption = "Q"
        s = 0
    End If
End Sub
'//++++++++++++++ End tmrFlyIn Code +++++++++++++++

'//++++++++++++++ tmrScroll Code ++++++++++++++++++
Private Sub tmrScroll_Timer()
    'scroll txtScroll
    If txtScroll.Top + txtScroll.Height < picScroll.Top Then 'picScroll.Top
        txtScroll.Top = picScroll.Height
    Else
        txtScroll.Top = txtScroll.Top - 25
    End If
End Sub
'//+++++++++++++ End tmrScroll code +++++++++++++++

'//++++++++++ txtScroll Code ++++++++++++++++++++++
Private Sub txtScroll_GotFocus()
    cmdOK.SetFocus
    'Don't let the text box get focus, although
    'the text box is locked it looks bad to see
    'a cursor in the text box as it scrolls up
End Sub
'//+++++++++++++ End txtScroll code +++++++++++++++

'//+++++++++++++ Form OnTop Code
Private Property Let OnTop(Setting As Boolean)
    If Setting Then
        'make this form topmost
        SetWindowPos hwnd, HWND_TOPMOST, _
        0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Else
        'Make this form non-topmost
        SetWindowPos hwnd, HWND_NOTOPMOST, _
        0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
    
    mbOnTop = Setting
End Property

Private Property Get OnTop() As Boolean
    'Return the private variable set in property Let
    OnTop = mbOnTop
End Property

