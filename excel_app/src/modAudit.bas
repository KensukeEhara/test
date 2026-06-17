Attribute VB_Name = "modAudit"
Option Explicit

Public Function GetWindowsIdentity() As String
    Dim domainName As String
    Dim userName As String

    domainName = Environ$("USERDOMAIN")
    userName = Environ$("USERNAME")

    If userName = "" Then userName = Application.UserName

    If domainName <> "" And userName <> "" Then
        GetWindowsIdentity = domainName & "\\" & userName
    Else
        GetWindowsIdentity = userName
    End If
End Function

Public Sub AppendAuditLog(ByVal assetNumber As String, ByVal fieldName As String, ByVal oldValue As String, ByVal newValue As String, ByVal reason As String)
    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_AUDIT)
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = Application.UserName
    ws.Cells(nextRow, 3).Value = GetWindowsIdentity()
    ws.Cells(nextRow, 4).Value = assetNumber
    ws.Cells(nextRow, 5).Value = fieldName
    ws.Cells(nextRow, 6).Value = oldValue
    ws.Cells(nextRow, 7).Value = newValue
    ws.Cells(nextRow, 8).Value = reason
End Sub
