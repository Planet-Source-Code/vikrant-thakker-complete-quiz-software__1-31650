VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Quiz Pro    by    AnaSys Softwares"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   10395
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   195
      Left            =   11760
      TabIndex        =   12
      Top             =   8400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2160
      ScaleHeight     =   375
      ScaleWidth      =   6975
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5760
      Width           =   6975
      Begin VB.TextBox txtScroll 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   6960
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         Top             =   0
         Width           =   200
      End
   End
   Begin VB.Timer tmrScroll 
      Interval        =   100
      Left            =   480
      Top             =   4680
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   615
      Left            =   30
      TabIndex        =   13
      Top             =   60
      Width           =   10335
   End
   Begin VB.Label lblAbout 
      BackColor       =   &H000000FF&
      Caption         =   " About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label lblPc 
      BackColor       =   &H000000FF&
      Caption         =   " Enter Price Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   4650
      TabIndex        =   6
      Top             =   2430
      Width           =   2325
   End
   Begin VB.Label lblExit 
      BackColor       =   &H000000FF&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label lblEnter 
      BackColor       =   &H000000FF&
      Caption         =   " Enter Questions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label lblPlay 
      BackColor       =   &H000000FF&
      Caption         =   " Play Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   1860
      Width           =   2295
   End
   Begin VB.Label lblEnter1 
      BackColor       =   &H0000FF00&
      Caption         =   " Enter Questions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   3000
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblPlay1 
      BackColor       =   &H0000FF00&
      Caption         =   " Play Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   1860
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblExit1 
      BackColor       =   &H0000FF00&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label lblPc1 
      BackColor       =   &H0000FF00&
      Caption         =   " Enter Price Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4650
      TabIndex        =   7
      Top             =   2430
      Width           =   2325
   End
   Begin VB.Label lblAbout1 
      BackColor       =   &H0000FF00&
      Caption         =   " About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   3600
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private DefaultDir As String
Dim ScrollText As Integer
Private Sub Form_Load()

Call HideDesktop
Call HideTaskBar
lblPlay1.Visible = False
'lblView1.Visible = False
lblEnter1.Visible = False
lblExit1.Visible = False

'//This is part of the scrolling text code
ScrollText = FreeFile
Open App.Path & "\Scroll.txt" For Input As ScrollText
txtScroll = Input(LOF(ScrollText), ScrollText)
Close #ScrollText
    
    HideALL 'call sub HideAll
    PositionAll 'call sub PositionAll
   ' fraFront.Visible = True 'Make fraFront visible
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblPlay1.Visible = False
'lblView1.Visible = False
lblEnter1.Visible = False
lblAbout1.Visible = False
lblPc1.Visible = False
lblExit1.Visible = False
lblPlay.Visible = True
'lblView.Visible = True
lblAbout.Visible = True
lblEnter.Visible = True
lblPc.Visible = True
lblExit.Visible = True

End Sub

Private Sub Form_Terminate()
Call ShowDesktop
Call ShowTaskBar
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call ShowDesktop
Call ShowTaskBar
End Sub

Private Sub lblEnter_Click()
Load frmMastQuest
frmMastQuest.Show
Unload Me
End Sub
Private Sub lblEnter1_Click()
Load frmMastQuest
frmMastQuest.Show
Unload Me
End Sub

Private Sub lblEnter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEnter1.Visible = True
lblEnter.Visible = False
End Sub

Private Sub lblExit_Click()
frmMain.Hide
frmInfo.Show
End Sub
Private Sub lblExit1_Click()
frmMain.Hide
frmInfo.Show
End Sub
Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit1.Visible = True
lblExit.Visible = False
End Sub

Private Sub lblPlay_Click()
Load frmName
frmName.Show
Unload Me
End Sub
Private Sub lblPlay1_Click()
Load frmName
frmName.Show
Unload Me
End Sub

Private Sub lblPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblPlay1.Visible = True
lblPlay.Visible = False
End Sub
Private Sub lblView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'lblView1.Visible = True
'lblView.Visible = False
End Sub
Private Sub lblpc_Click()
Load frmPc
frmPc.Show
Unload Me
End Sub
Private Sub lblpc1_Click()
Load frmPc
frmPc.Show
Unload Me
End Sub
Private Sub lblPc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblPc1.Visible = True
lblPc.Visible = False
End Sub

Private Sub lblAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAbout1.Visible = True
lblAbout.Visible = False
End Sub


Private Sub lblAbout_Click()
Load frmAbout
frmAbout.Show
Unload Me
End Sub
Private Sub lblAbout1_Click()
Load frmAbout
frmAbout.Show
Unload Me
End Sub


'//Other Sub Code Starts Here
Private Sub HideALL()
Dim ctl As Control

For Each ctl In Me.Controls
    If (TypeOf ctl Is Frame) Then
        ctl.Visible = False
    End If
Next
End Sub

Private Sub PositionAll()
Dim ctl As Control

For Each ctl In Me.Controls
    If (TypeOf ctl Is Frame) Then
        ctl.Top = 1080 '840
        ctl.Left = 120
        ctl.BorderStyle = 0
        'ctl.Height = 4215
        ctl.Width = 6375
    End If
Next
End Sub
'//Other Sub Code Ends Here


'//Text Box Code Starts Here
'The scrolling textboxes can't get focus, not for a major
'reason it just looks sloppy

Private Sub txtScroll_GotFocus()
   ' cmdFront.SetFocus
'Text1.SetFocus
End Sub
'//Text Box Code Starts Here


'//Scrolling Text Code Starts Here
Private Sub tmrScroll_Timer()
Dim Temp As Long
'This resizez text box to fit the text
Temp = Len(txtScroll.Text)
Temp = Temp * 105
txtScroll.Width = Temp
txtScroll.Left = txtScroll.Left - 40

If (txtScroll.Left + txtScroll.Width) < Picture1.Left Then
    txtScroll.Left = Picture1.ScaleWidth
End If
End Sub
'//Scrolling Text Code Ends Here


