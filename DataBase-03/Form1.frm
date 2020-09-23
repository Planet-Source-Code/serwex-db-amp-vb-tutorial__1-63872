VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   13
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Update 
      Caption         =   "Update"
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Delete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Add 
      Caption         =   "Add"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text5 
      DataField       =   "password"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      DataField       =   "credit"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      DataField       =   "lname"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      DataField       =   "fname"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      DataField       =   "id"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Data Data1 
      Caption         =   "DataBase Simple operations"
      Connect         =   "Access"
      DatabaseName    =   "C:\CivilDataBank\Educational\VBClass\DataBase-1\bbb.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "users"
      Top             =   2400
      Width           =   4380
   End
   Begin VB.Label Label5 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Credit:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Last Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "First Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   975
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
' Tutorial  : 3#18
'----------------------------------------------------------------------

'There are 18 step by step article/applications about how to contact
'databases (M.S. Access [*.MDB] as default)... unfortunately there is no
'comments inside , but so simple to understand!!!

'These samples are useful for beginners in VB...(Any feedbacks welcome)

Option Explicit

Private Sub Add_Click()
Data1.Recordset.AddNew
Add.Enabled = False
Delete.Enabled = False
Update.Enabled = True
Cancel.Enabled = True

End Sub

Private Sub Cancel_Click()
Data1.Recordset.CancelUpdate
Add.Enabled = True
Delete.Enabled = True
Update.Enabled = False
Cancel.Enabled = False

End Sub

Private Sub Delete_Click()
Data1.Recordset.Delete
Data1.Recordset.MoveFirst
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\bbb.mdb"
Add.Enabled = True
Delete.Enabled = True
Update.Enabled = False
Cancel.Enabled = False


End Sub

Private Sub Update_Click()
Data1.Recordset.Update
Add.Enabled = True
Delete.Enabled = True
Update.Enabled = False
Cancel.Enabled = False

End Sub
