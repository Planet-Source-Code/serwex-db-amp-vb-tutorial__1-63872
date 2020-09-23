VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Find Example..."
   ClientHeight    =   1290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   ScaleHeight     =   1290
   ScaleWidth      =   3825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "..."
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text2 
      DataField       =   "fname"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      DataField       =   "id"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1080
      TabIndex        =   2
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
      Top             =   840
      Width           =   3615
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
' Tutorial  : 8#18
'----------------------------------------------------------------------

'There are 18 step by step article/applications about how to contact
'databases (M.S. Access [*.MDB] as default)... unfortunately there is no
'comments inside , but so simple to understand!!!

'These samples are useful for beginners in VB...(Any feedbacks welcome)

Option Explicit
Dim varName As Variant
Dim strBkMark As String

Private Sub Command1_Click()

varName = InputBox("Enter search Criteria for Author Name:", "Find ID", "*102*")
If varName = "" Then
    Exit Sub
Else
    varName = "'" & varName & "'" 'String Syntax in SQL
End If

With Data1.Recordset
    strBkMark = .Bookmark
    .FindFirst "id Like" & varName
    If .NoMatch Then
        .Bookmark = strBkMark
        MsgBox "Unable to find!", vbInformation
    End If
End With
End Sub

Private Sub Command2_Click()

varName = InputBox("Enter search Criteria for Author Name:", "Find ID", "*bb*")
If varName = "" Then
    Exit Sub
Else
    varName = "'" & varName & "'" 'String Syntax in SQL
End If

With Data1.Recordset
    strBkMark = .Bookmark
    .FindFirst "fname Like" & varName
    If .NoMatch Then
        .Bookmark = strBkMark
        MsgBox "Unable to find!", vbInformation
    End If
End With

End Sub

Private Sub Command3_Click()

With Data1.Recordset
    strBkMark = .Bookmark
    .FindNext "id Like" & varName
    If .NoMatch Then
        .Bookmark = strBkMark
        MsgBox "Unable to find!", vbInformation
    End If
End With

End Sub

Private Sub Command4_Click()


With Data1.Recordset
    strBkMark = .Bookmark
    .FindNext "fname Like" & varName
    If .NoMatch Then
        .Bookmark = strBkMark
        MsgBox "Unable to find!", vbInformation
    End If
End With


End Sub

Private Sub Form_Load()

Data1.DatabaseName = App.Path & "\bbb.mdb"

End Sub

'//////////////////////////////////\\\\\\
'************ Super Star -Ultra- Extra Examples ************
'For Numeric Fields Use Find method like bellow:
'   .FindFirst "id=" & varName
'and For Date Fields Use Find method like bellow:
'   .FindFirst "date1=" & varName
'and Format the input by:
'    varDateStart= "#" & varDateStart & "#" 'Date Syntax in SQL
