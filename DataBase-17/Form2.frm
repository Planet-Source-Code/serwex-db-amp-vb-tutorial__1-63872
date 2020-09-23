VERSION 5.00
Begin VB.Form frmResults 
   Caption         =   "Form2"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5940
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3960
   ScaleWidth      =   5940
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form2.frx":0000
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "frmResults"
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

Private Sub Form_Load()
With Text1
    .Left = 0
    .Top = 0
    .Width = Me.ScaleWidth
    .Height = Me.ScaleHeight
'    .SetFocus
    .ZOrder
End With

End Sub

Private Sub mnuDBEngine_Click()
ShowForm frmDBEngine
End Sub
