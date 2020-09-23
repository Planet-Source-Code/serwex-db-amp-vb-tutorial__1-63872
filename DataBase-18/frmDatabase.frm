VERSION 5.00
Begin VB.Form frmDatabase 
   Caption         =   "Form1"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3495
   ScaleWidth      =   3930
   Begin VB.CommandButton Command8 
      Caption         =   "Attach Table"
      Height          =   615
      Left            =   2040
      TabIndex        =   7
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Modify Table"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "CreateTable"
      Height          =   615
      Left            =   2040
      TabIndex        =   5
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "TableDef"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CreateProperty"
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Execute"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OpenRS"
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Properties"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmDatabase"
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
' Tutorial  : 18#18
'----------------------------------------------------------------------

'There are 18 step by step article/applications about how to contact
'databases (M.S. Access [*.MDB] as default)... unfortunately there is no
'comments inside , but so simple to understand!!!

'These samples are useful for beginners in VB...(Any feedbacks welcome)

Private Sub Command1_Click()
Dim ws As Workspace
Dim db As Database
Dim strDBName As String
strDBName = App.Path & "\dao1.mdb"
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(strDBName)
DisplayResults GetProperties(db), "DB Properties"


End Sub

Private Sub Command2_Click()
Dim ws As Workspace
Dim db As Database
Dim rsTable As Recordset
Dim rsDynaset As Recordset
Dim rsSnapshot As Recordset
Dim rsForwardOnly As Recordset
Dim rsTemp As Recordset
'
Dim strDBName As String
Dim strRSTable As String
Dim strRSDynaset As String
Dim strRSSnapshot As String
Dim strRSForwardOnly As String
Dim strMsg As String
'
strDBName = App.Path & "\amir.mdb"
strRSTable = "main"
strRSDynaset = "driverinfo"
strRSSnapshot = "insuranceinfo"
strRSForwardOnly = "main"
'
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(strDBName)
'
With db
    Set rsTable = .OpenRecordset(strRSTable, dbOpenTable)
    Set rsDynaset = .OpenRecordset(strRSDynaset, dbOpenDynaset)
    Set rsSnapshot = .OpenRecordset(strRSSnapshot, dbOpenSnapshot)
    Set rsForwardOnly = .OpenRecordset(strRSForwardOnly, dbOpenForwardOnly)
End With

For Each rsTemp In db.Recordsets
    strMsg = strMsg & GetProperties(rsTemp) & vbCrLf
Next
DisplayResults strMsg, "DB OpenRecordset"
'
For Each rsTemp In db.Recordsets
    rsTemp.Close
    Set rsTemp = Nothing
Next
db.Close
Set db = Nothing


End Sub

Private Sub Command3_Click()
    Dim ws As Workspace
    Dim db As Database
    '
    Dim strDBName As String
    Dim strSql As String
    Dim IngRecords As Long
    '
    strDBName = App.Path & "\amir.mdb"
    strSql = "delete * from main where filenum='101'"
    IngRecords = 0
    '
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(strDBName)
    '
    With db
        .Execute strSql, dbFailOnError
        IngRecords = .RecordsAffected
    End With
    '
    DisplayResults "Records Affected = " & CStr(IngRecords), "DB Execute"
    db.Close
    ws.Close
    Set db = Nothing
    Set ws = Nothing
End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim ws As Workspace
Dim db As Database
Dim pr As Property
'
Dim strDBName As String
Dim strUDPName As String
Dim intUDPType As String
Dim varUDPValue As Variant
Dim strMsg As String
'
strDBName = App.Path & "\amir.mdb"
'
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(strDBName)
'
strUDPName = "DBAdmin"
intUDPType = vbVariant
varUDPValue = "D.B. Guru"
'
With db
    .Properties.Delete strUDPName
    Set pr = .CreateProperty(strUDPName, intUDPType, varUDPValue)
    .Properties.Append pr
End With
'
strUDPName = "Programmer"
intUDPType = vbVariant
varUDPValue = "D.B. Coder"
'
With db
    .Properties.Delete strUDPName
    Set pr = .CreateProperty(strUDPName)
    pr.Type = intUDPType
    pr.Value = varUDPValue
    .Properties.Append pr
End With
'
DisplayResults GetProperties(db), "DB CreateProperties"
'
db.Close
ws.Close
Set db = Nothing
Set ws = Nothing
Set pr = Nothing

End Sub

Private Sub Command5_Click()
On Error Resume Next
'
Dim ws As Workspace
Dim db As Database
Dim td As TableDef
Dim pr As Property
'
Dim strDBName As String
Dim strTDName As String
Dim strMsg As String
'
strDBName = App.Path & "\amir.mdb"
strTDName = "NewTable"
'
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(strDBName)
'
strMsg = ""
For Each td In db.TableDefs
    strMsg = strMsg & GetProperties(td)
    strMsg = strMsg & vbCrLf
Next
'
DisplayResults strMsg, "DB TableDefs"
'
db.Close
ws.Close
Set pr = Nothing
Set td = Nothing
Set db = Nothing
Set ws = Nothing

End Sub

Private Sub Command6_Click()
On Error Resume Next
Dim ws As Workspace
Dim db As Database
Dim td As TableDef
Dim fl As Field
Dim pr As Property
'
Dim strDBName As String
Dim strTDName As String
Dim strFLName As String
Dim intFLType As String
Dim strMsg As String
'
strDBName = App.Path & "\aaa.mdb"
strTDName = "NewTable"
strFLName = "NewField"
intFLType = dbText
'
Kill strDBName
'
Set ws = DBEngine.Workspaces(0)
Set db = ws.CreateDatabase(strDBName, dbLangGeneral, dbVersion30)
'
'Create a new Table
Set td = db.CreateTableDef(strTDName)
'
'Create a new field in table
Set fl = td.CreateField(strFLName, intFLType)
'
td.Fields.Append fl
db.TableDefs.Append td
'
DisplayResults GetProperties(td), "DB CreateTableDef"
'
db.Close
ws.Close
Set pr = Nothing
Set td = Nothing
Set db = Nothing
Set ws = Nothing

End Sub

Private Sub Command7_Click()
On Error Resume Next
Dim ws As Workspace
Dim db As Database
Dim td As TableDef
Dim fl As Field
'
Dim strDBName As String
Dim strTDName As String
Dim strFLName As String
Dim intFLType As Integer
Dim strMsg As String
'
strDBName = App.Path & "\aaa.mdb"
strTDName = "NewTable"
strFLName = "FollowDate"
intFLType = dbDate
'
Command6_Click
'
Set ws = DBEngine.Workspaces(0)
Set db = OpenDatabase(strDBName)
Set td = db.TableDefs(strTDName)
'
Set fl = td.CreateField(strFLName, intFLType)
td.Fields.Append fl
'Make list of Fields
strMsg = "Appended Field:" & vbCrLf
For Each fl In td.Fields
    strMsg = strMsg & vbTab & fl.Name & vbCrLf
Next
'Delete New field
td.Fields.Delete strFLName
'Make list again
strMsg = strMsg & "Deleted Field:" & vbCrLf
For Each fl In td.Fields
    strMsg = strMsg & vbTab & fl.Name & vbCrLf
Next
'
DisplayResults strMsg, "DB Table Modifications"
'
db.Close
ws.Close
Set fl = Nothing
Set td = Nothing
Set db = Nothing
Set ws = Nothing

End Sub

Private Sub Command8_Click()
MsgBox "Refer to pages 250-252"

End Sub
