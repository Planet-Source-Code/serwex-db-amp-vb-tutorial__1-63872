VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   2895
      Left            =   120
      OleObjectBlob   =   "Form1.frx":0014
      TabIndex        =   6
      Top             =   1800
      Width           =   7215
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Future-Write\Educational\VBClass\DataBase-15\bbb.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "personalinfo"
      Top             =   4560
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox Text3 
      DataField       =   "lname"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      DataField       =   "fname"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   480
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      DataField       =   "id"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   3735
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Future-Write\Educational\VBClass\DataBase-15\bbb.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "users"
      Top             =   1320
      Width           =   4935
   End
   Begin VB.Label Label3 
      Caption         =   "Last Name:"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Fisrtname:"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "ID:"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   1335
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
' Tutorial  : 15#18
'----------------------------------------------------------------------

'There are 18 step by step article/applications about how to contact
'databases (M.S. Access [*.MDB] as default)... unfortunately there is no
'comments inside , but so simple to understand!!!

'These samples are useful for beginners in VB...(Any feedbacks welcome)

Private Sub Data1_Reposition()
Dim strSql As String
Dim StrLink As String

StrLink = Text1.Text

strSql = "select * from personalinfo where id='" & StrLink & "'"
Data2.RecordSource = strSql
Data2.Refresh
DBGrid1.ReBind
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\bbb.mdb"
Data2.DatabaseName = App.Path & "\bbb.mdb"
Call Data1_Reposition
End Sub
