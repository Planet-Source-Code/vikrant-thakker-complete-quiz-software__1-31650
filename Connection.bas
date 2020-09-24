Attribute VB_Name = "Connection"
Public conn As ADODB.Connection
Public rsCust As ADODB.Recordset
Public rsPrice As ADODB.Recordset
Public rsQs As ADODB.Recordset
Public rsGujQs As ADODB.Recordset
Public rsCity As ADODB.Recordset
Public rsTemp As ADODB.Recordset
Public rsPc As ADODB.Recordset

Public Sub Main()
'On Error GoTo merr
Dim str1 As String
Set conn = New ADODB.Connection

str1 = "provider=microsoft.jet.oledb.3.51;Jet OLEDB:database password=imBillgates;data source="
str1 = str1 & App.Path & "\QuestBank.mdb"
'str1 = "provider=microsoft.jet.oledb.4.0;data source="
'str1 = str1 & App.Path & "\Acc2000.mdb"


conn.Open str1


Set rsCust = New ADODB.Recordset
rsCust.Open "CustDet", conn, adOpenStatic, adLockOptimistic

Set rsQs = New ADODB.Recordset
rsQs.Open "MastQuest", conn, adOpenStatic, adLockOptimistic

Set rsTemp = New ADODB.Recordset
rsTemp.Open "Temp", conn, adOpenStatic, adLockOptimistic


Set rsPrice = New ADODB.Recordset
rsPrice.Open "PriceMast", conn, adOpenStatic, adLockOptimistic

Set rsCity = New ADODB.Recordset
rsCity.Open "City", conn, adOpenStatic, adLockOptimistic

Set rsGujQs = New ADODB.Recordset
rsGujQs.Open "GujQuest", conn, adOpenStatic, adLockOptimistic

Set rsPc = New ADODB.Recordset
rsPc.Open "Price", conn, adOpenStatic, adLockOptimistic

'Load frmMain
'frmMain.Show

Load frmSplash
frmSplash.Show
Exit Sub
merr:
MsgBox err.Description
End Sub

