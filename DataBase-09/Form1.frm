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
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   1320
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
' Tutorial  : 9#18
'----------------------------------------------------------------------

'There are 18 step by step article/applications about how to contact
'databases (M.S. Access [*.MDB] as default)... unfortunately there is no
'comments inside , but so simple to understand!!!

'These samples are useful for beginners in VB...(Any feedbacks welcome)

Option Explicit

Private Sub Form_Activate()
Dim tb As TableDef
Dim fl As Field
Dim ix As Index
'///////////////
Data1.DatabaseName = App.Path & "\bbb.mdb"
Data1.Refresh
'///////////////
With Me.Text1
    .Text = "Table Info:" & vbCrLf
  For Each tb In Data1.Database.TableDefs
    .Text = .Text & vbTab & tb.Name & vbCrLf
    For Each fl In tb.Fields
        .Text = .Text & vbTab & vbTab & fl.Name & vbCrLf
    Next
  Next
  
On Error Resume Next
.Text = .Text & vbCrLf & "Index Info:" & vbCrLf
For Each tb In Data1.Database.TableDefs
    .Text = .Text & vbTab & tb.Name & vbCrLf
    For Each ix In tb.Indexes
        .Text = .Text & vbTab & vbTab & ix.Name & vbCrLf
        .Text = .Text & "[" & ix.Fields & "]" & vbCrLf
    Next
  Next
End With

End Sub

Private Sub Form_Resize()
With Me.Text1
    .Left = 0
    .Top = 0
    .Width = Me.ScaleWidth
    .Height = Me.ScaleHeight
End With

End Sub
