VERSION 5.00
Begin VB.Form frmName 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter your Name "
   ClientHeight    =   1155
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
Load frmMain
frmMain.Show
End Sub

Private Sub Form_Activate()
Call HideDesktop
Call HideTaskBar

txtName.SetFocus
End Sub

Private Sub Form_Load()
Call HideDesktop
End Sub

Private Sub OKButton_Click()
Load frmKbc
frmKbc.Show
'frmKbc.Show
Unload Me
End Sub
Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If

End Sub
