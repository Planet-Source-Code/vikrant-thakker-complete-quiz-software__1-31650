VERSION 5.00
Begin VB.Form frmMastQuest 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Question Entry"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdChangeLang 
      Caption         =   "Change Language"
      Height          =   315
      Left            =   8760
      TabIndex        =   35
      Top             =   60
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.OptionButton optGuj 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gujarati"
      Height          =   255
      Left            =   2580
      TabIndex        =   33
      Top             =   60
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.OptionButton optEng 
      BackColor       =   &H00FFFFFF&
      Caption         =   "English"
      Height          =   255
      Left            =   1080
      TabIndex        =   32
      Top             =   60
      Value           =   -1  'True
      Width           =   1155
   End
   Begin VB.Frame frameFind 
      Caption         =   "Enter the question to Find"
      Height          =   1515
      Left            =   180
      TabIndex        =   31
      Top             =   1740
      Visible         =   0   'False
      Width           =   10275
      Begin VB.TextBox txtFind 
         Height          =   315
         Left            =   540
         TabIndex        =   18
         Top             =   480
         Width           =   9255
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Height          =   315
         Left            =   4200
         TabIndex        =   19
         Top             =   900
         Width           =   855
      End
      Begin VB.CommandButton cmbCan 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   5220
         TabIndex        =   20
         Top             =   900
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdFi 
      Caption         =   "Enter the Question to Find"
      Height          =   555
      Left            =   8640
      TabIndex        =   17
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6900
      TabIndex        =   16
      Top             =   5940
      Width           =   915
   End
   Begin VB.Frame Framebutton 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   360
      TabIndex        =   21
      Top             =   5880
      Width           =   6495
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0080C0FF&
         Caption         =   "Add"
         Height          =   420
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdRemove 
         BackColor       =   &H0080C0FF&
         Caption         =   "Remove"
         Height          =   420
         Left            =   3690
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdPrev 
         BackColor       =   &H0080C0FF&
         Caption         =   "Previous"
         Height          =   420
         Left            =   4590
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H0080C0FF&
         Caption         =   "Next"
         Height          =   420
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H0080C0FF&
         Caption         =   "Save"
         Height          =   420
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H0080C0FF&
         Caption         =   "Modify"
         Height          =   420
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H0080C0FF&
         Caption         =   "Cancel"
         Height          =   420
         Left            =   2790
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   45
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   4935
      Left            =   150
      TabIndex        =   0
      Top             =   480
      Width           =   10455
      Begin VB.ComboBox cmbField 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         ItemData        =   "frmMastQuest.frx":0000
         Left            =   7440
         List            =   "frmMastQuest.frx":0028
         TabIndex        =   8
         Text            =   "[Select One]"
         Top             =   3960
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox cmbLevel 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         ItemData        =   "frmMastQuest.frx":00CF
         Left            =   570
         List            =   "frmMastQuest.frx":00E2
         TabIndex        =   7
         Text            =   "[Select One]"
         Top             =   3990
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox cmbAns 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         ItemData        =   "frmMastQuest.frx":00F5
         Left            =   4200
         List            =   "frmMastQuest.frx":0105
         TabIndex        =   6
         Text            =   "[Select One]"
         Top             =   3960
         Width           =   2415
      End
      Begin VB.TextBox txtD 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   5520
         TabIndex        =   5
         Top             =   2760
         Width           =   4095
      End
      Begin VB.TextBox txtC 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   2760
         Width           =   4095
      End
      Begin VB.TextBox txtB 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   5520
         TabIndex        =   3
         Top             =   2280
         Width           =   4095
      End
      Begin VB.TextBox txtA 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   2280
         Width           =   4095
      End
      Begin VB.TextBox txtQuest 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   600
         TabIndex        =   1
         Top             =   960
         Width           =   9255
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Question"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   600
         TabIndex        =   30
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
         Height          =   255
         Left            =   600
         TabIndex        =   29
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   7440
         TabIndex        =   28
         Top             =   3600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Difficulty Level"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   540
         TabIndex        =   27
         Top             =   3600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Correct Option"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4200
         TabIndex        =   26
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "D"
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
         Left            =   5160
         TabIndex        =   25
         Top             =   2790
         Width           =   375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "C"
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
         Left            =   600
         TabIndex        =   24
         Top             =   2820
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "B"
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
         Left            =   5160
         TabIndex        =   23
         Top             =   2310
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "A"
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
         Left            =   600
         TabIndex        =   22
         Top             =   2310
         Width           =   375
      End
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Language :"
      Height          =   315
      Left            =   180
      TabIndex        =   34
      Top             =   60
      Width           =   1275
   End
End
Attribute VB_Name = "frmMastQuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Boolean
Dim modi As Boolean
Private Sub cmbAns_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub
Private Sub cmbCan_Click()
txtFind.Text = ""
Frame1.Visible = True
Framebutton.Visible = True
frameFind.Visible = False
cmdClose.Visible = True
cmdFi.Visible = True
End Sub
'Private Sub cmbField_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    SendKeys "{TAB}"
'    KeyAscii = 0
'End If
'End Sub
'Private Sub cmbLevel_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    SendKeys "{TAB}"
'    KeyAscii = 0
'End If
'End Sub
Private Sub cmdChangeLang_Click()
flag = False
If frameFind.Visible = True Then
'MsgBox "Visible !"
    If optEng.Value = True Then
    Call optEng_Click
    optGuj.Value = True
    optEng.Value = False
    txtFind.Text = ""
    'Exit Sub
    ElseIf optEng.Value = False Then
    Call optGuj_Click
    optEng.Value = True
    optGuj.Value = False
    txtFind.Text = ""
    'Exit Sub
    End If

Exit Sub
End If

If Trim(txtQuest.Text) = "" Then
                                   ' Call showall
                                    If optEng.Value = True Then
                                    optGuj.Value = True
                                    optEng.Value = False
                                    Call optGuj_Click
                                    Exit Sub
            
                                    ElseIf optEng.Value = False Then
                                    optEng.Value = True
                                    optGuj.Value = False
                                    Call optEng_Click
                                    Exit Sub
                                    End If
'Call showall
End If

If cmdModify.Enabled = True Then

Call cmdCancel_Click

                                Call ClearAll
                                If optEng.Value = True Then
                                optGuj.Value = True
                                optEng.Value = False
                                Call optGuj_Click
                                    
                                ElseIf optEng.Value = False Then
                                optEng.Value = True
                                optGuj.Value = False
                                Call optEng_Click
                                End If


Exit Sub
End If


If Trim(txtQuest.Text) <> "" And cmdSave.Enabled = True Then
ans = MsgBox("Do you want to save this question", vbYesNo + vbQuestion)
    
        
If ans = vbNo Then

Call cmdCancel_Click

                                Call ClearAll
                                If optEng.Value = True Then
                                optGuj.Value = True
                                optEng.Value = False
                                Call optGuj_Click
                                Exit Sub
        
                                ElseIf optEng.Value = False Then
                                optEng.Value = True
                                optGuj.Value = False
                                Call optEng_Click
                                Exit Sub
                                End If
          
'Call showall
          
ElseIf ans = vbYes Then
                                Call cmdSave_Click
                    
                                    If flag = False Then
                                    Exit Sub
                                    End If
                
        If flag = True Then
        
                                If optEng.Value = True Then
                                optGuj.Value = True
                                optEng.Value = False
                                Call optGuj_Click
                                Call ClearAll
                                Exit Sub
                                ElseIf optEng.Value = False Then
                                optEng.Value = True
                                optGuj.Value = False
                                Call optEng_Click
                                Call ClearAll
                                Exit Sub
                                End If
        Call ClearAll
        
                 End If
    End If

End If
optEng.Enabled = False
optGuj.Enabled = False
End Sub
Private Sub ClearAll()
txtQuest.Text = ""
txtA.Text = ""
txtB.Text = ""
txtC.Text = ""
txtD.Text = ""
cmbAns.Text = "[Select One]"
'cmbLevel.Text = "[Select One]"
'cmbField.Text = "[Select One]"
End Sub
Private Sub cmdClose_Click()
Load frmMain
frmMain.Show
Unload Me
End Sub
Private Sub Command2_Click()
frmMastQuest.Show
End Sub
Private Sub Command3_Click()
frmQuiz.Show
End Sub
Private Sub Command4_Click()
frmQz.Show
End Sub
Private Sub cmdAdd_Click()
On Error GoTo aerr
modi = False
Call EnableAll
txtQuest.Enabled = True
txtA.Enabled = True
txtB.Enabled = True
txtC.Enabled = True
txtD.Enabled = True
cmbAns.Enabled = True
'cmbLevel.Enabled = True
'cmbField.Enabled = True

txtQuest.Text = ""
txtA.Text = ""
txtB.Text = ""
txtC.Text = ""
txtD.Text = ""
cmbAns.Text = "[Select One]"
'cmbLevel.Text = "[Select One]"
'cmbField.Text = "[Select One]"


'If optEng.Value = True Then
'rsQs.AddNew
'ElseIf optEng.Value = False Then
'rsGujQs.AddNew
'End If
cmdCancel.Enabled = True
cmdSave.Enabled = True
cmdModify.Enabled = False
cmdRemove.Enabled = False
cmdPrev.Enabled = False
cmdNext.Enabled = False
cmdAdd.Enabled = False
cmdClose.Enabled = False

txtQuest.SetFocus

Exit Sub
aerr:
MsgBox err.Description
End Sub
Private Sub cmdCancel_Click()
On Error GoTo cerr
modi = False
If optEng.Value = True Then
rsQs.CancelUpdate
ElseIf optEng.Value = False Then
rsGujQs.CancelUpdate
End If

Call DisableAll
txtQuest.Enabled = False
txtA.Enabled = False
txtB.Enabled = False
txtC.Enabled = False
txtD.Enabled = False
cmbAns.Enabled = False
'cmbLevel.Enabled = False
'cmbField.Enabled = False


cmdAdd.Enabled = True
cmdModify.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdRemove.Enabled = True
cmdCancel.Enabled = False
cmdSave.Enabled = False
cmdClose.Enabled = True

Call cmdPrev_Click

Exit Sub
cerr:
MsgBox err.Description
End Sub
Private Sub cmdFi_Click()
txtFind.Text = ""

frameFind.Visible = True
txtFind.SetFocus
Frame1.Visible = False
Framebutton.Visible = False
cmdClose.Visible = False
cmdFi.Visible = False
End Sub
Private Sub cmdFind_Click()
'On Error GoTo err
If optEng.Value = True Then
    If rsQs.EOF = False Or rsQs.BOF = False Then rsQs.MoveFirst
    Do While rsQs.EOF = False
    If UCase(txtFind.Text) = rsQs!Quest Then
    
    'cmbMonth.Visible = False
    'cmbDay.Visible = False
    'cmbYear.Visible = False
    'txtDt.Visible = True
    'txtDt.Enabled = False
    'cmdModify.Enabled = True
    
txtQuest.Text = rsQs("Quest")
txtA.Text = rsQs("A")
txtB.Text = rsQs("B")
txtC.Text = rsQs("C")
txtD.Text = rsQs("D")
cmbAns.Text = rsQs("Ans")
'cmbLevel.Text = rsQs("Level")
'cmbField.Text = rsQs("Field")

txtFind.Text = ""
        Frame1.Visible = True
        Framebutton.Visible = True
        cmdClose.Visible = True
        cmdFi.Visible = True
        frameFind.Visible = False
        cmdModify.Enabled = True
        cmdRemove.Enabled = True
        'txtSearch.SetFocus
                
        Exit Sub
    End If
    rsQs.MoveNext
Loop
If rsQs.EOF = True Then MsgBox "Record Not found !"
txtFind.Text = ""
txtFind.SetFocus
'Exit Sub
        
    '    End If
        
ElseIf optEng.Value = False Then
    If rsGujQs.EOF = False Or rsGujQs.BOF = False Then rsGujQs.MoveFirst
    Do While rsGujQs.EOF = False
    If UCase(txtFind.Text) = rsGujQs!Quest Then

txtQuest.Text = rsGujQs("Quest")
txtA.Text = rsGujQs("A")
txtB.Text = rsGujQs("B")
txtC.Text = rsGujQs("C")
txtD.Text = rsGujQs("D")
cmbAns.Text = rsGujQs("Ans")
'cmbLevel.Text = rsGujQs("Level")
'cmbField.Text = rsGujQs("Field")


txtFind.Text = ""
        Frame1.Visible = True
        Framebutton.Visible = True
        cmdClose.Visible = True
        cmdFi.Visible = True
        frameFind.Visible = False
        cmdModify.Enabled = True
        cmdRemove.Enabled = True
        'txtSearch.SetFocus
                
        Exit Sub
    End If
    rsGujQs.MoveNext
Loop
If rsGujQs.EOF = True Then MsgBox "Record Not found !"
txtFind.Text = ""
txtFind.SetFocus
'Exit Sub

End If
               
     Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub cmdModify_Click()
modi = True
Call EnableAll
txtQuest.SetFocus
cmdCancel.Enabled = True
cmdSave.Enabled = True
cmdPrev.Enabled = False
cmdNext.Enabled = False
cmdAdd.Enabled = False
cmdRemove.Enabled = False
cmdModify.Enabled = False
cmdClose.Enabled = False
End Sub
Private Sub cmdNext_Click()
On Error GoTo nerr
modi = False
Call DisableAll
cmdModify.Enabled = True

If optEng.Value = True Then
If rsQs.EOF = False Then rsQs.MoveNext
If rsQs.BOF = True And rsQs.EOF = True Then Exit Sub
If rsQs.EOF = True Then rsQs.MoveLast
'End If

ElseIf optEng.Value = False Then
If rsGujQs.EOF = False Then rsGujQs.MoveNext
If rsGujQs.BOF = True And rsGujQs.EOF = True Then Exit Sub
If rsGujQs.EOF = True Then rsGujQs.MoveLast
End If

showall

cmdRemove.Enabled = True
Exit Sub
nerr:
MsgBox err.Description
End Sub
Private Sub cmdPrev_Click()
On Error GoTo perr
modi = False
Call DisableAll
cmdModify.Enabled = True

If optEng.Value = True Then
If rsQs.BOF = True And rsQs.EOF = True Then Exit Sub
If rsQs.BOF = False Then rsQs.MovePrevious
If rsQs.BOF = True Then rsQs.MoveFirst
'End If

ElseIf optEng.Value = False Then
If rsGujQs.BOF = True And rsGujQs.EOF = True Then Exit Sub
If rsGujQs.BOF = False Then rsGujQs.MovePrevious
If rsGujQs.BOF = True Then rsGujQs.MoveFirst
End If

showall

cmdRemove.Enabled = True
Exit Sub
perr:
MsgBox err.Description
End Sub
Private Sub cmdRemove_Click()
On Error GoTo rerr
modi = False
Call DisableAll

If optEng.Value = True Then
If rsQs.BOF = True And rsQs.EOF = True Then Exit Sub
rsQs.Delete
rsQs.MoveNext
If rsQs.EOF = True Then rsQs.MovePrevious
'End If

ElseIf optEng.Value = False Then
If rsGujQs.BOF = True And rsGujQs.EOF = True Then Exit Sub
rsGujQs.Delete
rsGujQs.MoveNext
If rsGujQs.EOF = True Then rsGujQs.MovePrevious
End If

Call showall
Exit Sub
rerr:
MsgBox err.Description
End Sub
Private Sub cmdSave_Click()
On Error GoTo serr
flag = False
'MsgBox "flag = false"

If Trim(txtQuest.Text) = "" Then
MsgBox "Please enter the question !"
txtQuest.SetFocus
flag = False
    Exit Sub
End If

If modi = True Then
MsgBox "Update !"

rsQs!Quest = UCase(txtQuest.Text)
If txtA.Text <> "" Then rsQs!A = txtA.Text
If txtB.Text <> "" Then rsQs!B = txtB.Text
If txtC.Text <> "" Then rsQs!C = txtC.Text
If txtD.Text <> "" Then rsQs!d = txtD.Text
If cmbAns.Text <> "" Then rsQs!ans = cmbAns.Text
'If cmbLevel.Text <> "" Then rsQs!Level = cmbLevel.Text
'If cmbField.Text <> "" Then rsQs!Field = cmbField.Text
'If optEng = True Then rsQs!Language = "English"
rsQs.Update
'rsQs.Update
MsgBox "Updated !"
End If


If modi = False Then
If optEng.Value = True Then
'If rsQs.BOF = True Then MsgBox "File is empty"
If rsQs.BOF = False Then
rsQs.CancelUpdate
    rsQs.MoveFirst
    Do While rsQs.EOF = False
        If UCase(rsQs!Quest) = UCase(txtQuest) Then Exit Do
        rsQs.MoveNext
    Loop
    If rsQs.EOF = False Then
    MsgBox "Question already exists !"
    txtQuest.SetFocus
    Exit Sub
    End If
    If rsQs.EOF = True Then
       ' MsgBox "Question NOT EXIST !"
       rsQs.AddNew
 '       rsQs.MoveFirst
        End If
        End If
        
ElseIf optEng.Value = False Then
If rsGujQs.BOF = True Then MsgBox "File empty !"
If rsGujQs.BOF = False Then
    rsGujQs.MoveFirst
      
    Do While rsGujQs.EOF = False
        If UCase(rsGujQs!Quest) = UCase(txtQuest) Then Exit Do
        rsGujQs.MoveNext
    Loop
    If rsGujQs.EOF = False Then
    MsgBox "Question already exists !"
    txtQuest.SetFocus
    Exit Sub
    End If
    If rsGujQs.EOF = True Then
       ' MsgBox "Question NOT EXIST !"
   '     rsGujQs.AddNew
'        rsGujQs.MoveFirst
        End If
        End If
        End If
        End If
        








If Trim(txtA.Text) = "" Then
MsgBox "Please enter the option 'A'"
txtA.SetFocus
flag = False
Exit Sub
End If

If Trim(txtB.Text) = "" Then
MsgBox "Please enter the option 'B'"
txtB.SetFocus
flag = False
Exit Sub
End If

If Trim(txtC.Text = "") Then
MsgBox "Please enter the option 'C'"
txtC.SetFocus
flag = False
Exit Sub
End If

If Trim(txtD.Text = "") Then
MsgBox "Please enter the option 'D'"
txtD.SetFocus
flag = False
Exit Sub
End If


If Not (cmbAns.Text = "A" Or cmbAns.Text = "B" Or cmbAns.Text = "C" Or cmbAns.Text = "D") Then
MsgBox "Please select correct 'Answer' from the list !"
cmbAns.SetFocus
flag = False
Exit Sub
End If

'If Not (cmbLevel.Text = "1" Or cmbLevel.Text = "2" Or cmbLevel.Text = "3" Or cmbLevel.Text = "4" Or cmbLevel.Text = "5") Then
'MsgBox "Please select correct 'Level' from the list!"
'cmbLevel.SetFocus
'flag = False
'Exit Sub
'End If

'If Not (cmbField.Text = "Business" Or cmbField.Text = "Computers and Technology" Or cmbField.Text = "Entertainment" Or cmbField.Text = "Health" Or _
'cmbField.Text = "Home and Family" Or cmbField.Text = "Music" Or cmbField.Text = "Other" Or cmbField.Text = "Personal Finance" Or cmbField.Text = "Shopping" Or _
'cmbField.Text = "Small Business" Or cmbField.Text = "Sports and Outdoors" Or cmbField.Text = "Travel") Then

'MsgBox "Please select the 'Category' from List !"
'cmbField.Text = "[Select One]"
'cmbField.SetFocus
'flag = False
'Exit Sub
'End If


Call EnableAll
txtQuest.Enabled = False
txtA.Enabled = False
txtB.Enabled = False
txtC.Enabled = False
txtD.Enabled = False
cmbAns.Enabled = False
'cmbLevel.Enabled = False
'cmbField.Enabled = False



'If optEng.Value = True Then
'rsQs.AddNew
'ElseIf optEng.Value = False Then
'rsGujQs.AddNew
'End If


If optEng.Value = True And modi = False Then
'If modi = False Then
rsQs.CancelUpdate

rsQs.AddNew

rsQs!Quest = UCase(txtQuest.Text)
If txtA.Text <> "" Then rsQs!A = txtA.Text
If txtB.Text <> "" Then rsQs!B = txtB.Text
If txtC.Text <> "" Then rsQs!C = txtC.Text
If txtD.Text <> "" Then rsQs!d = txtD.Text
If cmbAns.Text <> "" Then rsQs!ans = cmbAns.Text
'If cmbLevel.Text <> "" Then rsQs!Level = cmbLevel.Text
'If cmbField.Text <> "" Then rsQs!Field = cmbField.Text
'If optEng = True Then rsQs!Language = "English"
rsQs.Update
'End If

ElseIf optEng.Value = False And modi = False Then
'If modi = False Then

rsQs.CancelUpdate

rsGujQs.AddNew
rsGujQs!Quest = UCase(txtQuest.Text)
If txtA.Text <> "" Then rsGujQs!A = txtA.Text
If txtB.Text <> "" Then rsGujQs!B = txtB.Text
If txtC.Text <> "" Then rsGujQs!C = txtC.Text
If txtD.Text <> "" Then rsGujQs!d = txtD.Text
If cmbAns.Text <> "" Then rsGujQs!ans = cmbAns.Text
'If cmbLevel.Text <> "" Then rsGujQs!Level = cmbLevel.Text
'If cmbField.Text <> "" Then rsGujQs!Field = cmbField.Text
'If optGuj = True Then rsGujQs!Language = "Gujarati"
rsGujQs.Update
End If



cmdSave.Enabled = False
cmdAdd.Enabled = True
cmdModify.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdCancel.Enabled = False
cmdRemove.Enabled = True
cmdClose.Enabled = True
flag = True
'MsgBox "flag = true"
Exit Sub
serr:
MsgBox "You have entered any invalid value !"
End Sub
Private Sub showall()
'On Error GoTo serr

txtQuest.Text = ""
txtA.Text = ""
txtB.Text = ""
txtC.Text = ""
txtD.Text = ""
cmbAns.Text = "[Select One]"
'cmbLevel.Text = "[Select One]"
'cmbField.Text = "[Select One]"


If optEng.Value = True Then
If rsQs.BOF = True And rsQs.EOF = True Then Exit Sub
If rsQs.BOF = False Or rsQs.EOF = False Then
    If Not rsQs!Quest = "" Then txtQuest.Text = rsQs!Quest
    If Not rsQs!A = "" Then txtA.Text = rsQs!A
    If Not rsQs!B = "" Then txtB.Text = rsQs!B
    If Not rsQs!C = "" Then txtC.Text = rsQs!C
    If Not rsQs!d = "" Then txtD.Text = rsQs!d
    If Not rsQs!ans = "" Then cmbAns.Text = rsQs!ans
   ' If Not rsQs!Level = "" Then cmbLevel.Text = rsQs!Level
 '   If Not rsQs!Field = "" Then cmbField.Text = rsQs!Field
           
End If
ElseIf optEng.Value = False Then
If rsGujQs.BOF = True And rsGujQs.EOF = True Then Exit Sub
If rsGujQs.BOF = False Or rsGujQs.EOF = False Then
    If Not rsGujQs!Quest = "" Then txtQuest.Text = rsGujQs!Quest
    If Not rsGujQs!A = "" Then txtA.Text = rsGujQs!A
    If Not rsGujQs!B = "" Then txtB.Text = rsGujQs!B
    If Not rsGujQs!C = "" Then txtC.Text = rsGujQs!C
    If Not rsGujQs!d = "" Then txtD.Text = rsGujQs!d
    If Not rsGujQs!ans = "" Then cmbAns.Text = rsGujQs!ans
'    If Not rsGujQs!Level = "" Then cmbLevel.Text = rsGujQs!Level
 '   If Not rsGujQs!Field = "" Then cmbField.Text = rsGujQs!Field
           
End If

End If
Exit Sub
'serr:
'MsgBox err.Description
End Sub


Private Sub EnableAll()
Frame1.Enabled = True
txtQuest.Enabled = True
txtA.Enabled = True
txtB.Enabled = True
txtC.Enabled = True
txtD.Enabled = True
cmbAns.Enabled = True
'cmbLevel.Enabled = True
'cmbField.Enabled = True
End Sub

Private Sub DisableAll()
Frame1.Enabled = False
txtQuest.Enabled = False
txtA.Enabled = False
txtB.Enabled = False
txtC.Enabled = False
txtD.Enabled = False
cmbAns.Enabled = False
'cmbLevel.Enabled = False
'cmbField.Enabled = False
End Sub
Private Sub Form_Load()
Call HideDesktop
Call HideTaskBar

modi = False
Call DisableAll
cmdCancel.Enabled = False
cmdRemove.Enabled = False
cmdModify.Enabled = False
cmdSave.Enabled = False

cmdAdd.Enabled = False
cmdPrev.Enabled = False
cmdNext.Enabled = False
cmdFi.Enabled = False


frameFind.Visible = False
cmbAns.Text = "[Select One]"
'cmbLevel.Text = "[Select One]"
'cmbField.Text = "[Select One]"
Call optEng_Click
End Sub

Private Sub Form_Terminate()
Call ShowDesktop
Call ShowTaskBar
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call ShowDesktop
Call ShowTaskBar
End Sub

Private Sub optEng_Click()
cmdAdd.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdFi.Enabled = True

txtQuest.Font.Name = "arial"
txtA.Font.Name = "arial"
txtB.Font.Name = "arial"
txtC.Font.Name = "arial"
txtD.Font.Name = "arial"
txtFind.Font.Name = "arial"

optGuj.Enabled = False
optEng.Enabled = False
End Sub

Private Sub optGuj_Click()
cmdAdd.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdFi.Enabled = True

txtQuest.Font.Name = "arial"
txtA.Font.Name = "arial"
txtB.Font.Name = "arial"
txtC.Font.Name = "arial"
txtD.Font.Name = "arial"
txtFind.Font.Name = "arial"

optEng.Enabled = False
optGuj.Enabled = False
End Sub
Private Sub txtA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub
Private Sub txtB_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub
Private Sub txtC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub
Private Sub txtFind_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtQuest_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If

End Sub

