VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   6210
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7935
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2400
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuDBEngine 
      Caption         =   "DBEngine"
   End
   Begin VB.Menu mnuWorkSpace 
      Caption         =   "WorkSpaces"
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "DataBase"
   End
End
Attribute VB_Name = "MDIForm1"
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

Private Sub mnuDatabase_Click()
ShowForm frmDatabase
End Sub

Private Sub mnuDBEngine_Click()
ShowForm frmDBEngine
End Sub

Private Sub mnuWorkSpace_Click()
ShowForm frmWorkspace

End Sub
