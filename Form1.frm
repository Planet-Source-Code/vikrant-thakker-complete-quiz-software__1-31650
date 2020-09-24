VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrScroll 
      Interval        =   100
      Left            =   2400
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6975
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   6975
      Begin VB.TextBox txtScroll 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         Top             =   120
         Width           =   200
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private DefaultDir As String
Dim ScrollText As Integer

Private Sub Form_Load()
    
'//This is part of the scrolling text code
ScrollText = FreeFile
Open App.Path & "\Scroll.txt" For Input As ScrollText
txtScroll = Input(LOF(ScrollText), ScrollText)
Close #ScrollText
    
    HideALL 'call sub HideAll
    PositionAll 'call sub PositionAll
   ' fraFront.Visible = True 'Make fraFront visible
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


