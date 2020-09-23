VERSION 5.00
Begin VB.Form frmDBEngine 
   Caption         =   "Form1"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4530
   ScaleWidth      =   7095
   Begin VB.CommandButton Command4 
      Caption         =   "Register DB"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Compact"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Repair"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Properties"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmDBEngine"
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
DisplayResults GetProperties(DBEngine), "DBEngine Properties"

End Sub

Private Sub Command2_Click()
On Error GoTo LocalErr
Dim strDBName As String
strDBName = GetDBFile("repair.mdb")
If strDBName <> "" Then
    DBEngine.RepairDatabase strDBName
    DisplayResults strDBName & " repaired", "DBEngine Repair"
End If
Exit Sub
LocalErr:
'\\\
End Sub

Private Sub Command3_Click()
Dim strOldDBName As String
Dim strNewDBName As String
Dim intEncrypt As Integer
Dim strVersion As String
Dim intVersion As Integer
CompactStart:
    strOldDBName = ""
    strNewDBName = ""
    strOldDBName = GetDBFile()
    If strOldDBName = "" Then
        Exit Sub
    End If
    strNewDBName = GetDBFile(strOldDBName)
    If strNewDBName = "" Then
        GoTo CompactStart
    End If
SetVersion:
    intVersion = 0
    strVersion = ""
    strVersion = InputBox("Select target version" & vbCrLf & _
    "1.1, 2.x, 3.x", "Select Version", "3.x")
    Select Case LCase(strVersion)
        Case "1.x": intVersion = dbVersion11
        Case "2.x": intVersion = dbVersion20
        Case "3.x": intVersion = dbVersion30
        Case Else
            MsgBox "Invalid Version!", vbCritical, "Version Error"
            GoTo SetVersion
    End Select
SetEncrypt:
    intEncrypt = MsgBox("Encrypt Target?", vbInformation + vbYesNo, "CompactDB")
    If intEncrypt = vbYes Then
        intEncrypt = dbEncrypt
    Else
        intEncrypt = dbDecrypt
    End If
RunCompact:
DBEngine.CompactDatabase strOldDBName, strNewDBName, _
dbLangGeneral, intVersion + intEncrypt
DisplayResults "Compact completed!", "DbEngine Compact"

End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim strDSN As String
Dim strDriver As String
Dim blnQuiet As Boolean
Dim strAttributes As String
Dim strDelimiter As String
strDelimiter = Chr(0)
strDSN = "TDPTest"
strDriver = "SQL Server"
blnQuiet = False

strAttributes = "SERVER=mca" & strDelimiter
strAttributes = strAttributes & "DATABASE=pubs" & strDelimiter
strAttributes = strAttributes & "DESCRIPTION=Sample Registration" & strDelimiter
DBEngine.RegisterDatabase strDSN, strDriver, blnQuiet, strAttributes


End Sub
