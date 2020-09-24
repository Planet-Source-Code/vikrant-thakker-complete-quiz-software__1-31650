VERSION 5.00
Begin VB.Form frmInfo 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AnaSys Softwares"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   4000
      Left            =   300
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   300
      Top             =   120
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hey Guys !!! Please dont forget to vote...........................  VIKRANT  THAKKER"
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
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   5280
      Width           =   7335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Email : vikrant_thakker@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   150
      TabIndex        =   2
      Top             =   3930
      Width           =   6855
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vikrant Thakker"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   0
      TabIndex        =   1
      Top             =   1830
      Width           =   6975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quiz  Software  developed by"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   6915
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Terminate()
Call ShowDesktop
Call ShowTaskBar
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call ShowDesktop
Call ShowTaskBar
End Sub

Private Sub Timer1_Timer()
lblName.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub

Private Sub Timer2_Timer()
 Call ShowDesktop
 Call ShowTaskBar
 Unload Me
 End
End Sub
