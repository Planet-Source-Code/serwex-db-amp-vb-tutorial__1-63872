VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Update"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Restore"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Data1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   3135
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   960
      Width           =   3060
   End
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
' Tutorial  : 10#18
'----------------------------------------------------------------------

'There are 18 step by step article/applications about how to contact
'databases (M.S. Access [*.MDB] as default)... unfortunately there is no
'comments inside , but so simple to understand!!!

'These samples are useful for beginners in VB...(Any feedbacks welcome)

Option Explicit


Public Sub OpenDB()
Dim cDBName As String
Dim cTblName As String
Dim bExclusive As Boolean
Dim bReadOnly As Boolean
'
cDBName = App.Path & "\bbb.mdb"
cTblName = "users"
bExclusive = True
bReadOnly = True
'
Data1.DatabaseName = cDBName
Data1.RecordSource = cTblName
Data1.Exclusive = bExclusive
Data1.Options = dbDenyWrite + dbReadOnly
Data1.ReadOnly = bReadOnly
';;;;;;;;;;;;;
'Data1.Options = 0
'Data1.ReadOnly = False
';;;;;;;;;;;;;
'
Data1.Refresh


End Sub

Private Sub Command1_Click()
Data1.UpdateControls ' Restore textbox values
End Sub

Private Sub Command2_Click()
Data1.UpdateRecord ' Write controls to dynaset
End Sub

Private Sub Data1_Reposition()
'MsgBox "Repositioning the pointer...", vbInformation
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
'MsgBox "Validating Data...", vbInformation

End Sub

Private Sub Form_Load()
OpenDB
BindControls
End Sub

Public Sub BindControls()
Dim cField1 As String
'
cField1 = "id"
Text1.DataField = cField1

End Sub
