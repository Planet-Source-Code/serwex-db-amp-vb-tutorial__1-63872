VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   2520
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   2040
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   13
      Mask            =   "(###)###-####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "Money:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Masked Edit:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Mixed:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "UpperCase:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Numbers (ver 2):"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Numbers (ver 1): "
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
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
' Tutorial  : 12#18
'----------------------------------------------------------------------

'There are 18 step by step article/applications about how to contact
'databases (M.S. Access [*.MDB] as default)... unfortunately there is no
'comments inside , but so simple to understand!!!

'These samples are useful for beginners in VB...(Any feedbacks welcome)

Option Explicit

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim strValid As String
strValid = "0123456789"
If InStr(strValid, Chr(KeyAscii)) = 0 Then
    KeyAscii = 0
End If

End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim strValid As String
strValid = "0123456789+-."
'
If KeyAscii > 26 Then
    If InStr(strValid, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub Text4_KeyPress(KeyAscii As Integer)
Dim strValid As String
strValid = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
'
KeyAscii = Asc(UCase(Chr(KeyAscii)))
'
If KeyAscii > 26 Then
    If InStr(strValid, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End If

End Sub
'Extra Informations from Pages 114-...-122
