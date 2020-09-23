VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Save Bookmark"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      DataField       =   "password"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1080
      TabIndex        =   9
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox Text4 
      DataField       =   "credit"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Top             =   1200
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      DataField       =   "lname"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   840
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      DataField       =   "fname"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   480
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      DataField       =   "id"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\CivilDataBank\Educational\VBClass\DataBase-7\bbb.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "users"
      Top             =   1920
      Width           =   4815
   End
   Begin VB.Label Label5 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Credit:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "L-Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "F-Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
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
' Tutorial  : 7#18
'----------------------------------------------------------------------

'There are 18 step by step article/applications about how to contact
'databases (M.S. Access [*.MDB] as default)... unfortunately there is no
'comments inside , but so simple to understand!!!

'These samples are useful for beginners in VB...(Any feedbacks welcome)

Option Explicit

Private Sub Command1_Click()
Static blnFlag As Boolean
Static strBookmark As String

With Data1.Recordset
    If blnFlag = False Then
        blnFlag = True
        Command1.Caption = "&Restore Bookmark"
        strBookmark = .Bookmark
        MsgBox "BookMark Saved Successfully!", vbInformation
    Else
        blnFlag = False
        Command1.Caption = "&Save Bookmark"
        .Bookmark = strBookmark
    End If
End With
End Sub

Private Sub Form_Load()

Data1.DatabaseName = App.Path & "\bbb.mdb"

End Sub

