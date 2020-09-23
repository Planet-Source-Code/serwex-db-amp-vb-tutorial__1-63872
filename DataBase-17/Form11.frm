VERSION 5.00
Begin VB.Form frmWorkspace 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton Command4 
      Caption         =   "Open DB"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Create DB"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create WS"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Properties"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmWorkspace"
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
' Tutorial  : 17#18
'----------------------------------------------------------------------

'There are 18 step by step article/applications about how to contact
'databases (M.S. Access [*.MDB] as default)... unfortunately there is no
'comments inside , but so simple to understand!!!

'These samples are useful for beginners in VB...(Any feedbacks welcome)

Private Sub Command1_Click()
Dim objWS As Workspace
Dim strMsg As String
For Each objWS In DBEngine.Workspaces
    strMsg = strMsg & GetProperties(objWS)
    strMsg = strMsg & vbCrLf
Next
DisplayResults strMsg, "WS Properties"
End Sub

Private Sub Command2_Click()
Dim ws As Workspace
Dim strWSName As String
Dim strWSUser As String
Dim strWSPassword As String
strWSName = "ws" & App.EXEName
strWSUser = "admin"
strWSPassword = ""
Set ws = DBEngine.CreateWorkspace(strWSName, strWSUser, strWSPassword)
DBEngine.Workspaces.Append ws
Command1_Click
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim dbOne As Database
Dim dbTwo As Database
Dim ws As Workspace
Dim dbTemp As Database
Dim strDBNameOne As String
Dim strDBNameTwo As String
Dim strWSName As String
Dim strWSUser As String
Dim strWSPassword As String
Dim strMsg As String

strDBNameOne = App.Path & "\dao1.mdb"
strDBNameTwo = App.Path & "\dao2.mdb"

strWSName = App.EXEName
strWSUser = "admin"
strWSPassword = ""
Kill strDBNameOne
Kill strDBNameTwo
Set ws = DBEngine.CreateWorkspace(strWSName, strWSUser, strWSPassword)
With ws
    Set dbOne = .CreateDatabase(strDBNameOne, dbLangGeneral, dbVersion30)
    Set dbOne = .CreateDatabase(strDBNameTwo, dbLangGeneral, dbVersion30)
    For Each dbTemp In .Databases
        strMsg = strMsg & "Name: " & dbTemp.Name & vbCrLf
    Next
    
    DisplayResults strMsg, "WS CreateDataBase"
    dbOne.Close
    dbTwo.Close
End With
Set dbOne = Nothing
Set dbTwo = Nothing
Set ws = Nothing

End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim wsRW As Workspace
Dim wsRO As Workspace
Dim wsTemp As Workspace
Dim dbRW As Database
Dim dbRO As Database
Dim dbTemp As Database
'
Dim strWSrwName As String
Dim strWSroName As String
Dim strDBName As String
Dim strWSUser As String
Dim strWSPassword As String
Dim strMsg As String
'
strWSrwName = "wsrw"
strWSroName = "wsro"
strWSUser = "admin"
strWSPassword = ""
strDBName = App.Path & "\bbb.mdb"
With DBEngine
    Set wsRW = .CreateWorkspace(strWSrwName, strWSUser, strWSPassword)
    Set wsRO = .CreateWorkspace(strWSroName, strWSUser, strWSPassword)
    .Workspaces.Append wsRW
    .Workspaces.Append wsRO
End With
Set dbRW = wsRW.OpenDatabase(strDBName)
Set dbRO = wsRW.OpenDatabase(strDBName, ReadOnly:=True)
Command1_Click
dbRW.Close
dbRO.Close
wsRW.Close
wsRO.Close

Set dbRW = Nothing
Set dbRO = Nothing
Set wsRW = Nothing
Set wsRO = Nothing



End Sub




