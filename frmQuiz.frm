VERSION 5.00
Begin VB.Form frmQuiz 
   BackColor       =   &H00000000&
   Caption         =   "Quiz"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "LOCK  KAR  DIYA  JAYE  !!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   2640
      TabIndex        =   12
      Top             =   4200
      Width           =   4095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   8175
      Begin VB.OptionButton Option4 
         Caption         =   "D"
         Height          =   375
         Left            =   4320
         TabIndex        =   11
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton Option3 
         Caption         =   "C"
         Height          =   375
         Left            =   4320
         TabIndex        =   10
         Top             =   480
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         Caption         =   "B"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "A"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblOptA 
         BackColor       =   &H00808080&
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
         Height          =   375
         Left            =   840
         TabIndex        =   7
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
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
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label lblOptC 
         BackColor       =   &H00808080&
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
         Height          =   375
         Left            =   4800
         TabIndex        =   5
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblOptD 
         BackColor       =   &H00808080&
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
         Height          =   375
         Left            =   4800
         TabIndex        =   4
         Top             =   960
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Question"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   8175
      Begin VB.TextBox txtQuest 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   7815
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Index           =   0
      Left            =   4080
      TabIndex        =   0
      Top             =   6120
      Width           =   1335
   End
End
Attribute VB_Name = "frmQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Unload Me
End Sub
