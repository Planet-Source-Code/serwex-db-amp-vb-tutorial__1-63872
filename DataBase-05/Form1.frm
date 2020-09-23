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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------
' DateTime  : 1/1/2006 11:51
' Author    : Shahin Noursalhi
' Contact   : admin@MixofTix.net
' Title     : DB & VB [tutorial]
' Purpose   : 18 step by step samples about contacting DB via VB!
' Tutorial  : 5#18
'----------------------------------------------------------------------

'There are 18 step by step article/applications about how to contact
'databases (M.S. Access [*.MDB] as default)... unfortunately there is no
'comments inside , but so simple to understand!!!

'These samples are useful for beginners in VB...(Any feedbacks welcome)

Option Explicit

Private Sub Form_Load()
Dim db As Database
Dim rs As Recordset
Dim rs2 As Recordset

Dim strDBname As String
Dim strRSname As String
Dim intRecs As Integer
Dim strFilter As String


strDBname = App.Path & "\bbb.mdb"
strRSname = "users"
strFilter = "id='101'"

Set db = DBEngine.OpenDatabase(strDBname)
Set rs = db.OpenRecordset(strRSname, dbOpenDynaset)

With rs
    .MoveLast
    intRecs = .RecordCount
End With

MsgBox "Record Counts: " & intRecs, vbInformation
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
rs.Filter = strFilter
Set rs2 = rs.OpenRecordset

With rs2
    .MoveLast
    intRecs = .RecordCount
End With
MsgBox "Record Counts: " & intRecs, vbInformation

End Sub
