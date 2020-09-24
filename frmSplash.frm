VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrWait 
      Interval        =   2000
      Left            =   120
      Top             =   1920
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   2160
      Picture         =   "frmSplash.frx":0000
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   1080
      Picture         =   "frmSplash.frx":0442
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright  © 2001 AnaSys Software"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2640
      Width           =   6135
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplash.frx":0884
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   5895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Windows 95/98/ME"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version 3.00"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quiz 2001"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   0
      Top             =   3600
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   4095
      Left            =   4320
      Top             =   -120
      Width           =   1935
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Call HideDesktop
Call HideTaskBar
PlaySound App.Path & "\Welcome.wav"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call ShowDesktop
Call ShowTaskBar
End Sub

Private Sub tmrWait_Timer()
    frmSplash.Hide
    frmMain.Show
    Unload frmSplash
End Sub
