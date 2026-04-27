Attribute VB_Name = "OrderDbCore"
Option Explicit

'=========================================================
' 受注票DB兼スケジュール管理表: コアロジック
' - 製番の生成/検証
' - DB書込先行の決定（最初の空行）
' - 工程日付の検証
'=========================================================

Public Enum ProcessDateValidationResult
    pvrOk = 0
    pvrStartOnly = 1
    pvrEndOnly = 2
    pvrStartAfterEnd = 3
    pvrOverDue = 4
End Enum

Public Function BuildSeibanCandidate(ByVal targetDate As Date, _
                                     ByVal officeCode As String, _
                                     ByVal serial4 As Long, _
                                     ByVal branch2 As Long) As String
    Dim yy As String

    If Not IsValidOfficeCode(officeCode) Then
        Err.Raise vbObjectError + 9001, "BuildSeibanCandidate", "営業所区分は B/E/Q のいずれかを指定してください。"
    End If

    yy = Right$(CStr(Year(targetDate)), 2)

    BuildSeibanCandidate = yy & "-" & UCase$(officeCode) & _
                           Format$(serial4, "0000") & "-" & _
                           Format$(branch2, "00")
End Function

Public Function IsValidSeiban(ByVal seiban As String) As Boolean
    Dim re As Object

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^\d{2}-[BEQ]\d{4}-\d{2}$"
    re.IgnoreCase = False
    re.Global = False

    IsValidSeiban = re.Test(Trim$(seiban))
End Function

Public Function FindFirstEmptyRow(ByVal ws As Worksheet, _
                                  Optional ByVal keyColumn As Long = 1, _
                                  Optional ByVal startRow As Long = 2) As Long
    Dim r As Long
    Dim lastRow As Long

    If startRow < 1 Then
        startRow = 1
    End If

    lastRow = ws.Cells(ws.Rows.Count, keyColumn).End(xlUp).Row
    If lastRow < startRow Then
        FindFirstEmptyRow = startRow
        Exit Function
    End If

    For r = startRow To lastRow
        If Trim$(CStr(ws.Cells(r, keyColumn).Value)) = vbNullString Then
            FindFirstEmptyRow = r
            Exit Function
        End If
    Next r

    FindFirstEmptyRow = lastRow + 1
End Function

Public Function ValidateProcessDates(ByVal startDate As Variant, _
                                     ByVal endDate As Variant, _
                                     ByVal dueDate As Variant, _
                                     ByRef message As String) As ProcessDateValidationResult
    Dim hasStart As Boolean
    Dim hasEnd As Boolean

    hasStart = IsDate(startDate)
    hasEnd = IsDate(endDate)

    If (Not hasStart) And (Not hasEnd) Then
        message = "開始日・終了日とも未入力です（登録可）。"
        ValidateProcessDates = pvrOk
        Exit Function
    End If

    If hasStart And (Not hasEnd) Then
        message = "開始日のみ入力されています。終了日を入力してください。"
        ValidateProcessDates = pvrStartOnly
        Exit Function
    End If

    If (Not hasStart) And hasEnd Then
        message = "終了日のみ入力されています。開始日を入力してください。"
        ValidateProcessDates = pvrEndOnly
        Exit Function
    End If

    If CDate(startDate) > CDate(endDate) Then
        message = "開始日が終了日より後です。"
        ValidateProcessDates = pvrStartAfterEnd
        Exit Function
    End If

    If IsDate(dueDate) Then
        If CDate(endDate) > CDate(dueDate) Then
            message = "工程終了日が納期を超過しています（警告）。"
            ValidateProcessDates = pvrOverDue
            Exit Function
        End If
    End If

    message = "工程日付は妥当です。"
    ValidateProcessDates = pvrOk
End Function

Public Function CanInputProcess(ByVal processName As String, _
                                ByVal hasKikai As Boolean, _
                                ByVal hasFirmware As Boolean, _
                                ByVal hasApplication As Boolean, _
                                ByVal hasShizai As Boolean, _
                                ByVal hasSeizou As Boolean, _
                                ByVal hasService As Boolean) As Boolean
    Select Case Trim$(processName)
        Case "機械設計"
            CanInputProcess = hasKikai
        Case "ソフト設計"
            CanInputProcess = (hasFirmware Or hasApplication)
        Case "資材"
            CanInputProcess = hasShizai
        Case "製造"
            CanInputProcess = (hasSeizou Or hasService)
        Case "デバッグ", "検査"
            CanInputProcess = True
        Case Else
            CanInputProcess = False
    End Select
End Function

Public Function IsValidOfficeCode(ByVal officeCode As String) As Boolean
    Dim v As String
    v = UCase$(Trim$(officeCode))
    IsValidOfficeCode = (v = "B" Or v = "E" Or v = "Q")
End Function
