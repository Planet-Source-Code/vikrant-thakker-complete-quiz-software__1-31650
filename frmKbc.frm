VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSChrt20.OCX"
Begin VB.Form frmKbc 
   BackColor       =   &H00000000&
   Caption         =   "Quiz Game"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      BackColor       =   &H00000000&
      Caption         =   "Life Line"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   8880
      TabIndex        =   49
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cmdHint 
         Caption         =   "Public"
         Height          =   375
         Left            =   1920
         TabIndex        =   52
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdSkip 
         Caption         =   "Skip"
         Height          =   375
         Left            =   1080
         TabIndex        =   51
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "50 - 50"
         Height          =   375
         Left            =   240
         TabIndex        =   50
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Height          =   555
      Left            =   8850
      TabIndex        =   47
      Top             =   1200
      Width           =   2985
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Custumer Name"
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
         Height          =   345
         Left            =   30
         TabIndex        =   48
         Top             =   120
         Width           =   2955
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   8820
      TabIndex        =   44
      Top             =   7800
      Width           =   3015
      Begin VB.Label lblTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "Total  Amount  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   30
         TabIndex        =   46
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label lblWon 
         BackColor       =   &H00000000&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
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
         Height          =   255
         Left            =   1980
         TabIndex        =   45
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Lock It  !!"
      Height          =   615
      Left            =   2910
      TabIndex        =   43
      Top             =   5040
      Width           =   3135
   End
   Begin VB.Frame Frame3 
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
      Left            =   360
      TabIndex        =   6
      Top             =   2310
      Width           =   8175
      Begin VB.OptionButton optA 
         Caption         =   "A"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   495
      End
      Begin VB.OptionButton optB 
         Caption         =   "B"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton optC 
         Caption         =   "C"
         Height          =   375
         Left            =   4320
         TabIndex        =   8
         Top             =   480
         Width           =   495
      End
      Begin VB.OptionButton optD 
         Caption         =   "D"
         Height          =   375
         Left            =   4320
         TabIndex        =   7
         Top             =   960
         Width           =   495
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
         TabIndex        =   14
         Top             =   960
         Width           =   3135
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
         TabIndex        =   13
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblOptB 
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
         TabIndex        =   12
         Top             =   960
         Width           =   3015
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
         TabIndex        =   11
         Top             =   480
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   1635
      Left            =   60
      TabIndex        =   4
      Top             =   -30
      Width           =   8775
      Begin VB.TextBox txtQuest 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
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
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   8115
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   6660
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Status"
      ForeColor       =   &H00FFFFFF&
      Height          =   6015
      Left            =   8850
      TabIndex        =   0
      Top             =   1770
      Width           =   3015
      Begin VB.Label lbl_1 
         BackColor       =   &H00000000&
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   2400
         TabIndex        =   42
         Top             =   840
         Width           =   585
      End
      Begin VB.Label lbl_1 
         BackColor       =   &H00000000&
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   2400
         TabIndex        =   41
         Top             =   1320
         Width           =   555
      End
      Begin VB.Label lbl_1 
         BackColor       =   &H00000000&
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   2400
         TabIndex        =   40
         Top             =   1800
         Width           =   555
      End
      Begin VB.Label lbl_1 
         BackColor       =   &H00000000&
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   2400
         TabIndex        =   39
         Top             =   2280
         Width           =   555
      End
      Begin VB.Label lbl_1 
         BackColor       =   &H00000000&
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   2400
         TabIndex        =   38
         Top             =   2760
         Width           =   555
      End
      Begin VB.Label lbl_1 
         BackColor       =   &H00000000&
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   37
         Top             =   3180
         Width           =   555
      End
      Begin VB.Label lbl_1 
         BackColor       =   &H00000000&
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   2400
         TabIndex        =   36
         Top             =   3720
         Width           =   555
      End
      Begin VB.Label lbl_1 
         BackColor       =   &H00000000&
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   35
         Top             =   4200
         Width           =   555
      End
      Begin VB.Label lbl_1 
         BackColor       =   &H00000000&
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   34
         Top             =   4740
         Width           =   555
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   300
         TabIndex        =   33
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   300
         TabIndex        =   32
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   300
         TabIndex        =   31
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   300
         TabIndex        =   30
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   300
         TabIndex        =   29
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   300
         TabIndex        =   28
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   300
         TabIndex        =   27
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   300
         TabIndex        =   26
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   25
         Top             =   4740
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "<<>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   1260
         TabIndex        =   24
         Top             =   4740
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "<<>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   1260
         TabIndex        =   23
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "<<>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   1260
         TabIndex        =   22
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "<<>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   1260
         TabIndex        =   21
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "<<>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   1260
         TabIndex        =   20
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "<<>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   1260
         TabIndex        =   19
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "<<>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   1260
         TabIndex        =   18
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "<<>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   1260
         TabIndex        =   17
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "<<>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   1260
         TabIndex        =   16
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "<<>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   1260
         TabIndex        =   15
         Top             =   5280
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   20
         X1              =   1260
         X2              =   600
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   19
         X1              =   2310
         X2              =   1650
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   18
         X1              =   1260
         X2              =   600
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   16
         X1              =   2310
         X2              =   1650
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   17
         X1              =   1260
         X2              =   600
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   15
         X1              =   1260
         X2              =   600
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   14
         X1              =   2310
         X2              =   1650
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   13
         X1              =   1260
         X2              =   600
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   12
         X1              =   2310
         X2              =   1650
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   11
         X1              =   2310
         X2              =   1650
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   4
         X1              =   1260
         X2              =   600
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   5
         X1              =   2310
         X2              =   1650
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   6
         X1              =   1260
         X2              =   600
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   7
         X1              =   2310
         X2              =   1650
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   8
         X1              =   1260
         X2              =   600
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   10
         X1              =   1260
         X2              =   600
         Y1              =   4860
         Y2              =   4860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   9
         X1              =   2310
         X2              =   1650
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   1260
         X2              =   600
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   2310
         X2              =   1650
         Y1              =   4860
         Y2              =   4860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   2310
         X2              =   1650
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Label lbl_1 
         BackColor       =   &H00000000&
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   2
         Top             =   5280
         Width           =   555
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   1
         Top             =   5280
         Width           =   375
      End
   End
   Begin VB.Frame frameGraph 
      BackColor       =   &H00C0C0C0&
      Height          =   6795
      Left            =   120
      TabIndex        =   53
      Top             =   1680
      Width           =   11655
      Begin VB.CommandButton Command5 
         Caption         =   "OK"
         Height          =   495
         Left            =   9600
         TabIndex        =   59
         Top             =   5640
         Width           =   1455
      End
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   5715
         Left            =   660
         OleObjectBlob   =   "frmKbc.frx":0000
         TabIndex        =   54
         Top             =   360
         Width           =   8055
      End
      Begin VB.Label lbl6 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   9240
         TabIndex        =   67
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "D"
         Height          =   375
         Left            =   9000
         TabIndex        =   66
         Top             =   4200
         Width           =   135
      End
      Begin VB.Label Label4 
         Caption         =   "C"
         Height          =   255
         Left            =   9000
         TabIndex        =   65
         Top             =   3360
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "B"
         Height          =   375
         Left            =   9000
         TabIndex        =   64
         Top             =   2520
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "A"
         Height          =   375
         Left            =   9000
         TabIndex        =   63
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label lbl4 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   9240
         TabIndex        =   62
         Top             =   4080
         Width           =   2295
      End
      Begin VB.Label lbl3 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   9240
         TabIndex        =   61
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label lbl2 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   9240
         TabIndex        =   60
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label lblD 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "44"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   6780
         TabIndex        =   58
         Top             =   6180
         Width           =   915
      End
      Begin VB.Label lblC 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "34"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   5400
         TabIndex        =   57
         Top             =   6180
         Width           =   855
      End
      Begin VB.Label lblB 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "24"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   4080
         TabIndex        =   56
         Top             =   6180
         Width           =   855
      End
      Begin VB.Label lblA 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   2700
         TabIndex        =   55
         Top             =   6180
         Width           =   855
      End
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   50
      Left            =   10680
      Picture         =   "frmKbc.frx":23B5
      Top             =   1320
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   49
      Left            =   8880
      Picture         =   "frmKbc.frx":A3EF
      Top             =   1320
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   48
      Left            =   7200
      Picture         =   "frmKbc.frx":12429
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   47
      Left            =   5400
      Picture         =   "frmKbc.frx":1A463
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   46
      Left            =   3600
      Picture         =   "frmKbc.frx":2249D
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   45
      Left            =   1800
      Picture         =   "frmKbc.frx":2A4D7
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   6
      Left            =   0
      Picture         =   "frmKbc.frx":32511
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   23
      Left            =   5400
      Picture         =   "frmKbc.frx":3A54B
      Top             =   1830
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   44
      Left            =   7170
      Picture         =   "frmKbc.frx":42585
      Top             =   1860
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   4
      Left            =   0
      Picture         =   "frmKbc.frx":4A5BF
      Top             =   1860
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   5
      Left            =   0
      Picture         =   "frmKbc.frx":525F9
      Top             =   540
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   43
      Left            =   0
      Top             =   1260
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   42
      Left            =   8970
      Picture         =   "frmKbc.frx":5A633
      Top             =   1830
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   41
      Left            =   10800
      Picture         =   "frmKbc.frx":6266D
      Top             =   30
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   40
      Left            =   10710
      Picture         =   "frmKbc.frx":6A6A7
      Top             =   1830
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   39
      Left            =   10800
      Picture         =   "frmKbc.frx":726E1
      Top             =   3180
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   38
      Left            =   10800
      Picture         =   "frmKbc.frx":7A71B
      Top             =   4530
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   37
      Left            =   10800
      Picture         =   "frmKbc.frx":82755
      Top             =   5880
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   36
      Left            =   10800
      Picture         =   "frmKbc.frx":8A78F
      Top             =   7290
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   35
      Left            =   9000
      Picture         =   "frmKbc.frx":927C9
      Top             =   30
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   34
      Left            =   9000
      Top             =   1830
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   33
      Left            =   9000
      Picture         =   "frmKbc.frx":9A803
      Top             =   3180
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   32
      Left            =   9000
      Picture         =   "frmKbc.frx":A283D
      Top             =   4530
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   31
      Left            =   9000
      Picture         =   "frmKbc.frx":AA877
      Top             =   5880
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   30
      Left            =   9000
      Picture         =   "frmKbc.frx":B28B1
      Top             =   7230
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   29
      Left            =   7080
      Picture         =   "frmKbc.frx":BA8EB
      Top             =   510
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   28
      Left            =   7200
      Picture         =   "frmKbc.frx":C2925
      Top             =   3210
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   27
      Left            =   7200
      Picture         =   "frmKbc.frx":CA95F
      Top             =   4560
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   26
      Left            =   7200
      Picture         =   "frmKbc.frx":D2999
      Top             =   5880
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   25
      Left            =   7200
      Picture         =   "frmKbc.frx":DA9D3
      Top             =   7230
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   24
      Left            =   5310
      Picture         =   "frmKbc.frx":E2A0D
      Top             =   480
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   22
      Left            =   5400
      Picture         =   "frmKbc.frx":EAA47
      Top             =   3240
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   21
      Left            =   5400
      Picture         =   "frmKbc.frx":F2A81
      Top             =   4560
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   20
      Left            =   5400
      Picture         =   "frmKbc.frx":FAABB
      Top             =   5910
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   19
      Left            =   5400
      Picture         =   "frmKbc.frx":102AF5
      Top             =   7260
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   18
      Left            =   3600
      Picture         =   "frmKbc.frx":10AB2F
      Top             =   540
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   17
      Left            =   3600
      Picture         =   "frmKbc.frx":112B69
      Top             =   1830
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   16
      Left            =   3600
      Picture         =   "frmKbc.frx":11ABA3
      Top             =   3210
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   15
      Left            =   3600
      Picture         =   "frmKbc.frx":122BDD
      Top             =   4560
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   14
      Left            =   3600
      Picture         =   "frmKbc.frx":12AC17
      Top             =   5910
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   13
      Left            =   3600
      Picture         =   "frmKbc.frx":132C51
      Top             =   7260
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   12
      Left            =   1800
      Picture         =   "frmKbc.frx":13AC8B
      Top             =   540
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   11
      Left            =   1800
      Picture         =   "frmKbc.frx":142CC5
      Top             =   1830
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   10
      Left            =   1800
      Picture         =   "frmKbc.frx":14ACFF
      Top             =   3210
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   9
      Left            =   1800
      Picture         =   "frmKbc.frx":152D39
      Top             =   4560
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   8
      Left            =   1800
      Picture         =   "frmKbc.frx":15AD73
      Top             =   5910
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   7
      Left            =   1800
      Picture         =   "frmKbc.frx":162DAD
      Top             =   7260
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   3
      Left            =   0
      Picture         =   "frmKbc.frx":16ADE7
      Top             =   3210
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   2
      Left            =   0
      Picture         =   "frmKbc.frx":172E21
      Top             =   4560
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   1
      Left            =   0
      Picture         =   "frmKbc.frx":17AE5B
      Top             =   5910
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1365
      Index           =   0
      Left            =   0
      Picture         =   "frmKbc.frx":182E95
      Top             =   7260
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1365
      Index           =   29
      Left            =   10020
      Top             =   1800
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1365
      Index           =   22
      Left            =   10050
      Top             =   3090
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1365
      Index           =   15
      Left            =   10050
      Top             =   4470
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1365
      Index           =   8
      Left            =   10050
      Top             =   5820
      Width           =   1800
   End
End
Attribute VB_Name = "frmKbc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Integer
Dim Ans1 As Boolean
Dim ans2 As Boolean
Dim ans3 As Boolean
Dim ans4 As Boolean
Dim ans5 As Boolean
Dim ans6 As Boolean
Dim ans7 As Boolean
Dim ans8 As Boolean
Dim ans9 As Boolean
Dim ans10 As Boolean
Dim d As Integer
Dim e As Integer
Dim f As Integer
Dim g As Integer
Dim ck As Boolean

Private Sub cmdHint_Click()



If rsQs.EOF = False Or rsQs.BOF = False Then rsQs.MoveFirst
Do While rsQs.EOF = False
    If txtQuest.Text = rsQs!Quest Then
       ' MsgBox "Record Found !"
    
'Command2.Enabled = False

    If rsQs!ans = "A" Then
        l = True
        m = False
        n = False
        o = False
    ElseIf rsQs!ans = "B" Then
          l = False
          m = True
          n = False
          o = False
    End If
    
    If rsQs!ans = "C" Then
          l = False
          m = False
          n = True
          o = False
          
    ElseIf rsQs!ans = "D" Then
          l = False
          m = False
          n = False
          o = True
    End If
    
   
       ' Exit Sub
    End If
    rsQs.MoveNext
Loop
'If rsQs.EOF = True Then MsgBox "Record Not found !"
'Exit Sub
'err:


lbl6.Caption = lblOptA.Caption
'lbl1.Item.Caption = lblOptA.Caption

lbl2.Caption = lblOptB.Caption
lbl3.Caption = lblOptC.Caption
lbl4.Caption = lblOptD.Caption

frameGraph.Visible = True
A = 25
Randomize
'B = Int((a - lowerbound + 1) * Rnd + lowerbound)
A = Int((A - 1 + 1) * Rnd + 1)
B = Int((A - 1 + 1) * Rnd + 1)
C = Int((A - 1 + 1) * Rnd + 1)
d = Int((A - 1 + 1) * Rnd + 1)

    
    e = 25 - A
    f = 25 - B
    g = 25 - C
    h = 25 - d
    
A = A + g
B = B + e
C = C + h
d = d + f

'MsgBox ("a=" & a)
'MsgBox ("b=" & b)
'MsgBox ("c=" & c)
'MsgBox ("d=" & d)


If l = True Then
A = A + 6
B = B - 2
C = C - 2
d = d - 2
End If


If m = True Then
A = A - 2
B = B + 6
C = C - 2
d = d - 2
End If

If n = True Then
A = A - 2
B = B - 2
C = C + 6
d = d - 2
End If

If o = True Then
A = A - 2
B = B - 2
C = C - 2
d = d + 6
End If


'MsgBox ("new a=" & a)
'MsgBox ("new b=" & b)
'MsgBox ("new c=" & c)
'MsgBox ("new d=" & d)




Dim X(1 To 5) As Variant
X(1) = " "
X(2) = A
X(3) = B
X(4) = C
X(5) = d
MSChart1.ChartData = X

lblA.Caption = A & "%"
lblB.Caption = B & "%"
lblC.Caption = C & "%"
lblD.Caption = d & "%"


End Sub

Private Sub cmdSkip_Click()
'flag = flag - 1
    lblOptA.Visible = True
    lblOptB.Visible = True
    lblOptC.Visible = True
    lblOptD.Visible = True
    
    optA.Visible = True
    optB.Visible = True
    optC.Visible = True
    optD.Visible = True
Call Throw

rsTemp.Delete
'flag = flag - 1
cmdSkip.Enabled = False
End Sub

Private Sub Command1_Click()
'End
If rsTemp.RecordCount = 0 Then
Load frmMain
frmMain.Show
Unload Me
End If

If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        For i = 1 To rsTemp.RecordCount Step 1
        rsTemp.Delete
        rsTemp.MoveNext
        Next
        Load frmMain
        frmMain.Show
        Unload Me
        End If

End Sub

Private Sub Command2_Click()
If rsQs.EOF = False Or rsQs.BOF = False Then rsQs.MoveFirst
Do While rsQs.EOF = False
    If txtQuest.Text = rsQs!Quest Then
       ' MsgBox "Record Found !"
    
Command2.Enabled = False

    If rsQs!ans = "A" Then
    lblOptA.Visible = True
    lblOptB.Visible = False
    lblOptC.Visible = False
    lblOptD.Visible = True
    
    lblOptA.Visible = True
    optB.Visible = False
    optC.Visible = False
    optD.Visible = True
    
        
    ElseIf rsQs!ans = "B" Then
    lblOptA.Visible = False
    lblOptB.Visible = True
    lblOptC.Visible = True
    lblOptD.Visible = False
    
    optA.Visible = False
    optB.Visible = True
    optC.Visible = True
    optD.Visible = False
    
    End If
    
    If rsQs!ans = "C" Then
    lblOptA.Visible = False
    lblOptB.Visible = True
    lblOptC.Visible = True
    lblOptD.Visible = False
        
    optA.Visible = False
    optB.Visible = True
    optC.Visible = True
    optD.Visible = False
    
    ElseIf rsQs!ans = "D" Then
    lblOptA.Visible = True
    lblOptB.Visible = False
    lblOptC.Visible = False
    lblOptD.Visible = True
    
    optA.Visible = True
    optB.Visible = False
    optC.Visible = False
    optD.Visible = True
    End If
    
   
        Exit Sub
    End If
    rsQs.MoveNext
Loop
If rsQs.EOF = True Then MsgBox "Record Not found !"
Exit Sub
'err:
'MsgBox err.Description
End Sub


Private Sub Command3_Click()


If optA.Value = False And optB.Value = False And optC.Value = False And optD.Value = False Then
MsgBox "Select the Option !"
Exit Sub
End If
Call check
'MsgBox (Flag)

If flag = 1 Then
lbl1(0).ForeColor = &H808080
lbl_1(0).ForeColor = &H808080
lblWon.Caption = rsPc!one
End If

If flag = 2 Then
For i = 0 To 1 Step 1
lbl1(i).ForeColor = &H808080
lbl_1(i).ForeColor = &H808080
Next
lblWon.Caption = rsPc!one + rsPc!two
End If

If flag = 3 Then
For i = 0 To 2 Step 1
lbl1(i).ForeColor = &H808080
lbl_1(i).ForeColor = &H808080
Next
lblWon.Caption = rsPc!one + rsPc!two + rsPc!three
End If

If flag = 4 Then
For i = 0 To 3 Step 1
lbl1(i).ForeColor = &H808080
lbl_1(i).ForeColor = &H808080
Next
lblWon.Caption = rsPc!one + rsPc!two + rsPc!three + rsPc!four
End If

If flag = 5 Then
For i = 0 To 4 Step 1
lbl1(i).ForeColor = &H808080
lbl_1(i).ForeColor = &H808080
Next
lblWon.Caption = rsPc!one + rsPc!two + rsPc!three + rsPc!four + rsPc!five
End If

If flag = 6 Then
For i = 0 To 5 Step 1
lbl1(i).ForeColor = &H808080
lbl_1(i).ForeColor = &H808080
Next
lblWon.Caption = rsPc!one + rsPc!two + rsPc!three + rsPc!four + rsPc!five + rsPc!six
End If

If flag = 7 Then
For i = 0 To 6 Step 1
lbl1(i).ForeColor = &H808080
lbl_1(i).ForeColor = &H808080
Next
lblWon.Caption = rsPc!one + rsPc!two + rsPc!three + rsPc!four + rsPc!five + rsPc!six + rsPc!seven
End If

If flag = 8 Then
For i = 0 To 7 Step 1
lbl1(i).ForeColor = &H808080
lbl_1(i).ForeColor = &H808080
Next
lblWon.Caption = rsPc!one + rsPc!two + rsPc!three + rsPc!four + rsPc!five + rsPc!six + rsPc!seven + rsPc!eight
End If

If flag = 9 Then
For i = 0 To 8 Step 1
lbl1(i).ForeColor = &H808080
lbl_1(i).ForeColor = &H808080
Next
lblWon.Caption = rsPc!one + rsPc!two + rsPc!three + rsPc!four + rsPc!five + rsPc!six + rsPc!seven + rsPc!eight + rsPc!nine
End If

If flag = 10 Then
For i = 0 To 9 Step 1
lbl1(i).ForeColor = &H808080
lbl_1(i).ForeColor = &H808080
Next
lblWon.Caption = rsPc!one + rsPc!two + rsPc!three + rsPc!four + rsPc!five + rsPc!six + rsPc!seven + rsPc!eight + rsPc!nine + rsPc!ten
 If rsTemp.BOF = False Or rsTemp.EOF = False Then
        rsTemp.MoveFirst
        For i = 1 To rsTemp.RecordCount Step 1
        rsTemp.Delete
        rsTemp.MoveNext
        Next
        End If
MsgBox ("Congratulations !! You win Rs :" + lblWon.Caption)
    optA.Enabled = False
    optB.Enabled = False
    optC.Enabled = False
    optD.Enabled = False
    Command3.Enabled = False
MsgBox ("Game Over !")
 Exit Sub
End If

If flag > 10 Then
For i = 0 To 9 Step 1
lbl1(i).ForeColor = &H808080
lbl_1(i).ForeColor = &H808080
Next
    optA.Enabled = False
    optB.Enabled = False
    optC.Enabled = False
    optD.Enabled = False
    Command3.Enabled = False
MsgBox ("GAME OVER !")
Exit Sub
End If

Call Throw

End Sub


Private Sub Command5_Click()
frameGraph.Visible = False
End Sub

Private Sub Form_Activate()
If rsPc!one = Empty Or rsPc!two = Empty Or rsPc!three = Empty Or rsPc!four = Empty Or rsPc!five = Empty Or rsPc!six = Empty Or rsPc!seven = Empty Or rsPc!eight = Empty Or rsPc!nine = Empty Or rsPc!ten = Empty Then
MsgBox "First Enter the price amount of all Questions !"
frmPc.Show
Unload Me
End If
End Sub

Private Sub Form_Load()
Call HideTaskBar
Call HideDesktop
Command3.Enabled = True

Command2.Enabled = True
frameGraph.Visible = False
optA.Enabled = True
optB.Enabled = True
optC.Enabled = True
optD.Enabled = True

If rsPc!one = Empty Or rsPc!two = Empty Or rsPc!three = Empty Or rsPc!four = Empty Or rsPc!five = Empty Or rsPc!six = Empty Or rsPc!seven = Empty Or rsPc!eight = Empty Or rsPc!nine = Empty Or rsPc!ten = Empty Then
MsgBox "First Enter the price amount of all Questions !"
frmPc.Show
Unload Me
End If



lblWon.Caption = "0"
ck = True
flag = 0
lblName.Caption = frmName.txtName.Text
Dim i As Integer
i = 0
For i = 0 To 9 Step 1
        lbl1(i).Caption = rsPrice!CorrAns
        lbl_1(i).Caption = rsPrice!PriceAmt
        lbl1(i).ForeColor = vbGreen
        lbl_1(i).ForeColor = vbGreen
Next


If rsPc.RecordCount > 0 Then
lbl1(0).Caption = "1"
lbl1(1).Caption = "2"
lbl1(2).Caption = "3"
lbl1(3).Caption = "4"
lbl1(4).Caption = "5"
lbl1(5).Caption = "6"
lbl1(6).Caption = "7"
lbl1(7).Caption = "8"
lbl1(8).Caption = "9"
lbl1(9).Caption = "10"


If rsPc!one <> "" Then lbl_1(0).Caption = rsPc!one
If rsPc!two <> "" Then lbl_1(1).Caption = rsPc!two
If rsPc!three <> "" Then lbl_1(2).Caption = rsPc!three
If rsPc!four <> "" Then lbl_1(3).Caption = rsPc!four
If rsPc!five <> "" Then lbl_1(4).Caption = rsPc!five
If rsPc!six <> "" Then lbl_1(5).Caption = rsPc!six
If rsPc!seven <> "" Then lbl_1(6).Caption = rsPc!seven
If rsPc!eight <> "" Then lbl_1(7).Caption = rsPc!eight
If rsPc!nine <> "" Then lbl_1(8).Caption = rsPc!nine
If rsPc!ten <> "" Then lbl_1(9).Caption = rsPc!ten
End If

End Sub

Private Sub Throw()

If ck = False Then
Exit Sub
End If

optA.Value = False
optB.Value = False
optC.Value = False
optD.Value = False
'MsgBox (flag)
'flag = flag + 1
repeat:
If rsQs.BOF = True And rsQs.EOF = True Then Exit Sub
If rsQs.BOF = False Then rsQs.MoveFirst

A = rsQs.RecordCount
Randomize
'B = Int((a - lowerbound + 1) * Rnd + lowerbound)
C = Int((A - 1 + 1) * Rnd + 1)
'C = Time
'e = B * c
'Randomize
'MsgBox (c)
rsQs.Move (C - 1)

If rsQs.EOF = True Then Exit Sub

If rsTemp.RecordCount = 0 Then
GoTo xyz
End If

If rsTemp.BOF = False Then
    rsTemp.MoveFirst
    Do While rsTemp.EOF = False
        If UCase(rsTemp!Quest) = UCase(rsQs!Quest) Then Exit Do
                 
        
        rsTemp.MoveNext
    Loop
    If rsTemp.EOF = False Then
    'MsgBox "Question already asked !"
    'rsQs.MoveNext
    GoTo repeat
    'Exit Sub
    End If
    If rsTemp.EOF = True Then
   '     MsgBox "City NOT EXIST"
xyz:
txtQuest.Text = rsQs!Quest
lblOptA.Caption = rsQs!A
lblOptB.Caption = rsQs!B
lblOptC.Caption = rsQs!C
lblOptD.Caption = rsQs!d

rsTemp.AddNew
rsTemp!Quest = txtQuest.Text
rsTemp.Update

 '       rsTemp.MoveFirst
        
    'MsgBox (rsTemp.RecordCount)
        Exit Sub
        End If
        End If



txtQuest.Text = rsQs!Quest
lblOptA.Caption = rsQs!A
lblOptB.Caption = rsQs!B
lblOptC.Caption = rsQs!C
lblOptD.Caption = rsQs!d

rsTemp.AddNew
rsTemp!Quest = txtQuest.Text
rsTemp.Update

End Sub

Private Sub check()
'On Error GoTo err
If rsQs.EOF = False Or rsQs.BOF = False Then rsQs.MoveFirst
Do While rsQs.EOF = False
    If txtQuest.Text = rsQs!Quest Then
       ' MsgBox "Record Found !"
    
    If rsQs!ans = "A" Then
    lblOptA.BackColor = vbRed
    'lblOptA.ForeColor = vbBlue
    lblOptA.FontBold = True
    lblOptA.ForeColor = vbCyan
    
    ElseIf rsQs!ans = "B" Then
    lblOptB.BackColor = vbRed
    'lblOptB.ForeColor = vbBlue
    lblOptB.FontBold = True
    lblOptB.ForeColor = vbCyan
    
    ElseIf rsQs!ans = "C" Then
    lblOptC.BackColor = vbRed
    'lblOptC.ForeColor = vbBlue
    lblOptC.FontBold = True
    lblOptC.ForeColor = vbCyan
    
    ElseIf rsQs!ans = "D" Then
    lblOptD.BackColor = vbRed
    'lblOptD.ForeColor = vbBlue
    lblOptD.FontBold = True
    lblOptD.ForeColor = vbCyan
    End If
    
    
    If optA.Value = True And rsQs!ans = "A" Then
    MsgBox "Correct Answer !"
    lblOptA.FontBold = False
    lblOptA.BackColor = &H808080
    lblOptA.ForeColor = vbWhite
    ck = True
    
    
    lblOptA.Visible = True
    lblOptB.Visible = True
    lblOptC.Visible = True
    lblOptD.Visible = True
    
    optA.Visible = True
    optB.Visible = True
    optC.Visible = True
    optD.Visible = True
    
    'flag = flag + 1
    flag = rsTemp.RecordCount
    End If
    
    If optA.Value = True And rsQs!ans <> "A" Then
    MsgBox "Wrong Answer !"
    ck = False
    End If
    
    If optB.Value = True And rsQs!ans = "B" Then
    MsgBox "Correct Answer !"
    lblOptB.FontBold = False
    lblOptB.BackColor = &H808080
    lblOptB.ForeColor = vbWhite
    ck = True
    'flag = flag + 1
    flag = rsTemp.RecordCount
    
    lblOptA.Visible = True
    lblOptB.Visible = True
    lblOptC.Visible = True
    lblOptD.Visible = True
    
    optA.Visible = True
    optB.Visible = True
    optC.Visible = True
    optD.Visible = True
    End If
    
    If optB.Value = True And rsQs!ans <> "B" Then
    MsgBox "Wrong Answer !"
    ck = False
    End If
    
    If optC.Value = True And rsQs!ans = "C" Then
    MsgBox "Correct Answer !"
    lblOptC.FontBold = False
    lblOptC.BackColor = &H808080
    lblOptC.ForeColor = vbWhite
    ck = True
    'flag = flag + 1
    flag = rsTemp.RecordCount
    lblOptA.Visible = True
    lblOptB.Visible = True
    lblOptC.Visible = True
    lblOptD.Visible = True
    
    optA.Visible = True
    optB.Visible = True
    optC.Visible = True
    optD.Visible = True
    End If
    
    If optC.Value = True And rsQs!ans <> "C" Then
    MsgBox "Wrong Answer !"
    ck = False
    End If
    
    If optD.Value = True And rsQs!ans = "D" Then
    MsgBox "Correct Answer !"
    lblOptD.FontBold = False
    lblOptD.BackColor = &H808080
    lblOptD.ForeColor = vbWhite
    ck = True
    'flag = flag + 1
    flag = rsTemp.RecordCount
    lblOptA.Visible = True
    lblOptB.Visible = True
    lblOptC.Visible = True
    lblOptD.Visible = True
    
    optA.Visible = True
    optB.Visible = True
    optC.Visible = True
    optD.Visible = True
    End If
    
    If optD.Value = True And rsQs!ans <> "D" Then
    MsgBox "Wrong Answer !"
    ck = False
    End If
    
       
    
    If ck = False Then
    MsgBox ("Congratulations !! You win Rs :" + lblWon.Caption)
    MsgBox "GAME OVER !!!"
        If rsTemp.RecordCount = 0 Then
        GoTo abc
        End If
        
        If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        For i = 1 To rsTemp.RecordCount Step 1
        rsTemp.Delete
        rsTemp.MoveNext
        Next
        End If
    'flag = 11
abc:
    optA.Enabled = False
    optB.Enabled = False
    optC.Enabled = False
    optD.Enabled = False
    Command3.Enabled = False
    Exit Sub
    End If
    
    
   
        Exit Sub
    End If
    rsQs.MoveNext
Loop
If rsQs.EOF = True Then MsgBox "Record Not found !"
Exit Sub
'err:
'MsgBox err.Description
End Sub

Private Sub Form_Paint()
Call Throw
If txtQuest.Text = "" Then
Call Throw
End If
lbl1(0).ForeColor = &H808080
lbl_1(0).ForeColor = &H808080
End Sub

Private Sub Form_Terminate()
Call ShowDesktop
Call ShowTaskBar
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call ShowDesktop
Call ShowTaskBar
End Sub

Private Sub lblOptA_Click()
optA.Value = True
End Sub

Private Sub lblOptB_Click()
optB.Value = True
End Sub

Private Sub lblOptC_Click()
optC.Value = True
End Sub

Private Sub lblOptD_Click()
optD.Value = True
End Sub

