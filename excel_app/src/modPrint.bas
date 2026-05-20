Attribute VB_Name = "modPrint"
Option Explicit

Public Sub PrintByAssetNumber(ByVal assetNumber As String)
    Dim ledgerWs As Worksheet
    Dim printWs As Worksheet
    Dim foundRow As Range

    If assetNumber = "" Then Err.Raise vbObjectError + 3001, , "資産番号を指定してください。"

    Set ledgerWs = ThisWorkbook.Worksheets(SHEET_LEDGER)
    Set printWs = ThisWorkbook.Worksheets(SHEET_PRINT)

    Set foundRow = ledgerWs.Columns(1).Find(What:=assetNumber, LookIn:=xlValues, LookAt:=xlWhole)
    If foundRow Is Nothing Then
        Err.Raise vbObjectError + 3002, , "対象資産番号が見つかりません。"
    End If

    printWs.Range("B2").Value = COMPANY_NAME
    printWs.Range("B3").Value = "資産番号"
    printWs.Range("C3").Value = assetNumber

    printWs.PrintOut
End Sub
