Attribute VB_Name = "modMasters"
Option Explicit

Public Function MasterCodeExists(ByVal sheetName As String, ByVal codeValue As String) As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets(sheetName)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If CStr(ws.Cells(i, 1).Value) = codeValue Then
            MasterCodeExists = True
            Exit Function
        End If
    Next i
End Function

Public Function DeprMethodExists(ByVal methodName As String) As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_DEPR_MASTER)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If CStr(ws.Cells(i, 1).Value) = methodName Then
            DeprMethodExists = True
            Exit Function
        End If
    Next i
End Function
