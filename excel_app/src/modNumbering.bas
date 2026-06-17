Attribute VB_Name = "modNumbering"
Option Explicit

Public Function GenerateAssetNumber(ByVal locationCode As String, ByVal typeCode As String, ByVal acquiredDate As Date) As String
    Dim yy As String
    Dim keyValue As String
    Dim serialValue As Long

    yy = Right$(CStr(Year(acquiredDate)), 2)
    keyValue = locationCode & typeCode & yy
    serialValue = NextSequence(keyValue)

    GenerateAssetNumber = keyValue & Format$(serialValue, "000")
End Function

Private Function NextSequence(ByVal keyValue As String) As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_SEQ)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If CStr(ws.Cells(i, 1).Value) = keyValue Then
            NextSequence = CLng(ws.Cells(i, 2).Value) + 1
            If NextSequence > 999 Then
                Err.Raise vbObjectError + 1001, , "連番上限999に達しました。"
            End If
            ws.Cells(i, 2).Value = NextSequence
            Exit Function
        End If
    Next i

    NextSequence = 1
    ws.Cells(lastRow + 1, 1).Value = keyValue
    ws.Cells(lastRow + 1, 2).Value = NextSequence
End Function
