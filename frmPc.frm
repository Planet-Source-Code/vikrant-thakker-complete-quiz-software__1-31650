VERSION 5.00
Begin VB.Form frmPc 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Enter Price Amount"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
      Height          =   405
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3780
      Width           =   855
   End
   Begin VB.Frame Framebutton 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   3480
      TabIndex        =   20
      Top             =   3720
      Width           =   2235
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save"
         Height          =   420
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   45
         Width           =   870
      End
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   495
      Left            =   6840
      TabIndex        =   9
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   495
      Left            =   3720
      TabIndex        =   7
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lbl10 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
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
      Left            =   6840
      TabIndex        =   19
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl9 
      BackStyle       =   0  'Transparent
      Caption         =   "9"
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
      Left            =   5280
      TabIndex        =   18
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "8"
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
      Left            =   3720
      TabIndex        =   17
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl7 
      BackStyle       =   0  'Transparent
      Caption         =   "7"
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
      Left            =   2160
      TabIndex        =   16
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lbl6 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
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
      Left            =   600
      TabIndex        =   15
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lbl5 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
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
      Left            =   6840
      TabIndex        =   14
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lbl4 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
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
      Left            =   5280
      TabIndex        =   13
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lbl3 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
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
      Left            =   3720
      TabIndex        =   12
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
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
      Left            =   2160
      TabIndex        =   11
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "frmPc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdClose_Click()
Load frmMain
frmMain.Show
Unload Me
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose.SetFocus
End Sub

Private Sub cmdSave_Click()
On Error GoTo err:

If Trim(Text1.Text) = "" Then
    Cancel = True
MsgBox "Enter the value in 1 !"
Text1.SetFocus
Exit Sub
End If

If Trim(Text2.Text) = "" Then
    Cancel = True
MsgBox "Enter the value in 2 !"
Text2.SetFocus
Exit Sub
End If
If Trim(Text3.Text) = "" Then
    Cancel = True
MsgBox "Enter the value in 3!"
Text3.SetFocus
Exit Sub
End If
If Trim(Text4.Text) = "" Then
    Cancel = True
MsgBox "Enter the value in 4!"
Text4.SetFocus
Exit Sub
End If
If Trim(Text5.Text) = "" Then
    Cancel = True
MsgBox "Enter the value in 5!"
Text5.SetFocus
Exit Sub
End If
If Trim(Text6.Text) = "" Then
    Cancel = True
MsgBox "Enter the value in 6!"
Text6.SetFocus
Exit Sub
End If
If Trim(Text7.Text) = "" Then
    Cancel = True
MsgBox "Enter the value in 7!"
Text7.SetFocus
Exit Sub
End If
If Trim(Text8.Text) = "" Then
    Cancel = True
MsgBox "Enter the value in 8!"
Text8.SetFocus
Exit Sub
End If
If Trim(Text9.Text) = "" Then
    Cancel = True
MsgBox "Enter the value in 9!"
Text9.SetFocus
Exit Sub
End If
If Trim(Text10.Text) = "" Then
    Cancel = True
MsgBox "Enter the value in 10!"
Text10.SetFocus
Exit Sub
End If






If rsPc.RecordCount = 0 Then
MsgBox "File is empty !"
rsPc.AddNew
If Not Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Then
rsPc!one = Text1.Text
rsPc!two = Text2.Text
rsPc!three = Text3.Text
rsPc!four = Text4.Text
rsPc!five = Text5.Text
rsPc!six = Text6.Text
rsPc!seven = Text7.Text
rsPc!eight = Text8.Text
rsPc!nine = Text9.Text
rsPc!ten = Text10.Text
rsPc.Update
'ElseIf Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Then
'MsgBox "No Field should remain empty !"
End If
End If

If rsPc.RecordCount > 0 Then
If Not Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Then
rsPc!one = Text1.Text
rsPc!two = Text2.Text
rsPc!three = Text3.Text
rsPc!four = Text4.Text
rsPc!five = Text5.Text
rsPc!six = Text6.Text
rsPc!seven = Text7.Text
rsPc!eight = Text8.Text
rsPc!nine = Text9.Text
rsPc!ten = Text10.Text
rsPc.Update
End If
End If
Exit Sub
err:
MsgBox "Make sure that you have not entered any Invalid Value !"
End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdSave.SetFocus
End Sub

Private Sub Form_Load()
Call HideDesktop
Call HideTaskBar

If rsPc.BOF = True And rsPc.EOF = True Then
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
End If

'If rsPc.EOF = False Or rsPc.BOF = False Then
If rsPc.RecordCount > 0 Then
If Not rsPc!one = "0" Then Text1.Text = rsPc!one
If Not rsPc!two = "0" Then Text2.Text = rsPc!two
If Not rsPc!three = "0" Then Text3.Text = rsPc!three
If Not rsPc!four = "0" Then Text4.Text = rsPc!four
If Not rsPc!five = "0" Then Text5.Text = rsPc!five
If Not rsPc!six = "0" Then Text6.Text = rsPc!six
If Not rsPc!seven = "0" Then Text7.Text = rsPc!seven
If Not rsPc!eight = "0" Then Text8.Text = rsPc!eight
If Not rsPc!nine = "0" Then Text9.Text = rsPc!nine
If Not rsPc!ten = "0" Then Text10.Text = rsPc!ten
End If



If rsPc.BOF = False Or rsPc.EOF = False Then
If rsPc!one = "0" Then Text1.Text = ""
If rsPc!two = "0" Then Text2.Text = ""
If rsPc!three = "0" Then Text3.Text = ""
If rsPc!four = "0" Then Text4.Text = ""
If rsPc!five = "0" Then Text5.Text = ""
If rsPc!six = "0" Then Text6.Text = ""
If rsPc!seven = "0" Then Text7.Text = ""
If rsPc!eight = "0" Then Text8.Text = ""
If rsPc!nine = "0" Then Text9.Text = ""
If rsPc!ten = "0" Then Text10.Text = ""
End If


End Sub


Private Sub Form_Unload(Cancel As Integer)
Call ShowDesktop
Call ShowTaskBar
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Trim(Text1.Text) = "" Then
    Cancel = True
MsgBox "This field cannot be empty !"
Text1.SetFocus
Exit Sub
End If

If Trim(Text1.Text) < 1 Then
Cancel = True
MsgBox "The value of this field cannot be smaller than 1 !"
Exit Sub
End If

End Sub

Private Sub Text2_Validate(Cancel As Boolean)
If Trim(Text2.Text) = "" Or Trim(Text2.Text) < 1 Then
    Cancel = True
MsgBox "This field cannot be empty !"
Text2.SetFocus
Exit Sub
End If

If Trim(Text2.Text) < 1 Then
Cancel = True
MsgBox "The value of this field cannot be smaller than 1 !"
Exit Sub
End If


End Sub

Private Sub Text3_Validate(Cancel As Boolean)
If Trim(Text3.Text) = "" Then
    Cancel = True
MsgBox "This field cannot be empty !"
Text3.SetFocus
Exit Sub
End If

If Trim(Text3.Text) < 1 Then
Cancel = True
MsgBox "The value of this field cannot be smaller than 1 !"
Exit Sub
End If


End Sub

Private Sub Text4_Validate(Cancel As Boolean)
If Trim(Text4.Text) = "" Then
    Cancel = True
MsgBox "This field cannot be empty !"
Text4.SetFocus
Exit Sub
End If

If Trim(Text4.Text) < 1 Then
Cancel = True
MsgBox "The value of this field cannot be smaller than 1 !"
Exit Sub
End If


End Sub

Private Sub Text5_Validate(Cancel As Boolean)
If Trim(Text5.Text) = "" Then
    Cancel = True
MsgBox "This field cannot be empty !"
Text5.SetFocus
Exit Sub
End If

If Trim(Text5.Text) < 1 Then
Cancel = True
MsgBox "The value of this field cannot be smaller than 1 !"
Exit Sub
End If


End Sub

Private Sub Text6_Validate(Cancel As Boolean)
If Trim(Text6.Text) = "" Then
    Cancel = True
MsgBox "This field cannot be empty !"
Text6.SetFocus
Exit Sub
End If

If Trim(Text6.Text) < 1 Then
Cancel = True
MsgBox "The value of this field cannot be smaller than 1 !"
Exit Sub
End If


End Sub

Private Sub Text7_Validate(Cancel As Boolean)
If Trim(Text7.Text) = "" Then
    Cancel = True
MsgBox "This field cannot be empty !"
Text7.SetFocus
Exit Sub
End If

If Trim(Text7.Text) < 1 Then
Cancel = True
MsgBox "The value of this field cannot be smaller than 1 !"
Exit Sub
End If


End Sub

Private Sub Text8_Validate(Cancel As Boolean)
If Trim(Text8.Text) = "" Then
    Cancel = True
MsgBox "This field cannot be empty !"
Text8.SetFocus
Exit Sub
End If

If Trim(Text8.Text) < 1 Then
Cancel = True
MsgBox "The value of this field cannot be smaller than 1 !"
Exit Sub
End If


End Sub

Private Sub Text9_Validate(Cancel As Boolean)
If Trim(Text9.Text) = "" Then
    Cancel = True
MsgBox "This field cannot be empty !"
Text9.SetFocus
Exit Sub
End If

If Trim(Text9.Text) < 1 Then
Cancel = True
MsgBox "The value of this field cannot be smaller than 1 !"
Exit Sub
End If


End Sub

Private Sub Text10_Validate(Cancel As Boolean)
If Trim(Text10.Text) = "" Then
    Cancel = True
MsgBox "This field cannot be empty !"
Text10.SetFocus
Exit Sub
End If

If Trim(Text10.Text) < 1 Then
Cancel = True
MsgBox "The value of this field cannot be smaller than 1 !"
Exit Sub
End If

End Sub

