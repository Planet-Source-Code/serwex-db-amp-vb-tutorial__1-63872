Attribute VB_Name = "Module1"
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

Public Function ShowType(TypeCode As Variant) As String
Dim strReturn As String
Select Case TypeCode
    Case vbEmpty
        strReturn = "Empty"
    Case vbNull
        strReturn = "Null"
    Case vbInteger
        strReturn = "Integer"
    Case vbLong
        strReturn = "Long"
    Case vbSingle
        strReturn = "Single"
    Case vbDouble
        strReturn = "Double"
    Case vbCurrency
        strReturn = "Currency"
    Case vbDate
        strReturn = "Date"
    Case vbString
        strReturn = "String"
    Case vbObject
        strReturn = "Object"
    Case vbError
        strReturn = "Error"
    Case vbVariant
        strReturn = "Variant"
    Case vbDataObject
        strReturn = "DataObject"
    Case vbDecimal
        strReturn = "Decimal"
    Case vbByte
        strReturn = "Byte"
    Case vbArray
        strReturn = "Array"
    Case Else
        strReturn = "[" & CStr(TypeCode) & "]"
End Select
ShowType = strReturn
End Function

Public Function GetProperties(objDAOItem As Object) As String
On Error GoTo LocalErr
Dim objItem As Object
Dim strReturn As String
Screen.ActiveForm.MousePointer = vbHourglass
strReturn = ""
For Each objItem In objDAOItem.Properties
    strReturn = strReturn & objItem.Name
    strReturn = strReturn & " = "
    If objItem.Name = "BookMark" Then
        strReturn = strReturn & "?"
    Else
        strReturn = strReturn & objItem.Value
    End If
    strReturn = strReturn & " { "
    strReturn = strReturn & ShowType(objItem.Type)
    strReturn = strReturn & " } " & vbCrLf
Next
GetProperties = strReturn
Screen.ActiveForm.MousePointer = vbNormal
Exit Function
LocalErr:
strReturn = strReturn & "<err>"
Resume Next
End Function

Public Function GetDBFile(Optional DefaultName As String = "") As String
On Error GoTo LocalErr
With MDIForm1.cd1
    .FileName = DefaultFileName
    .Filter = "*.mdb|*.mdb"
    .ShowOpen
    GetDBFile = .FileName
End With
Exit Function
LocalErr:
GetDBFile = ""
End Function

Public Sub DisplayResults(Data As String, Optional Title As String)
If IsMissing(Title) Then
    Title = ""
End If
With frmResults
    .Caption = Title
    .Text1.Text = Data
    .Show
End With

End Sub

Public Sub ShowForm(frmMe As Form)
Dim frm As Form
For Each frm In Forms
    If frm.Name <> "MDIForm1" Then
    Unload frm
    End If
Next
With frmMe
    .Left = 0
    .Top = 0
    .Width = MDIForm1.ScaleWidth / 2
    .Height = MDIForm1.ScaleHeight
End With
With frmResults
    .Left = MDIForm1.ScaleWidth / 2
    .Top = 0
    .Width = MDIForm1.ScaleWidth / 2
    .Height = MDIForm1.ScaleHeight
    .Text1.Text = ""
    .Caption = "Results for " & frmMe.Name
End With
End Sub
