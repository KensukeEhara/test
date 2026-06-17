Attribute VB_Name = "modAssetEntry"
Option Explicit

Public Sub InitializeWorkbook()
    SetupLedgerHeaders
    SetupAuditHeaders
    SetupSequenceHeaders
End Sub

Private Sub SetupLedgerHeaders()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_LEDGER)

    ws.Range("A1:Q1").Value = Array("資産番号", "資産種別コード", "購入orリース", "設置場所コード", "名称", "型式", "取得年月日", "取得額", "耐用年数", "償却方法", "償却率", "管理番号", "備考", "状態", "最終更新日時", "最終更新者", "最終更新WindowsID")
End Sub

Private Sub SetupAuditHeaders()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_AUDIT)

    ws.Range("A1:H1").Value = Array("変更日時", "編集者名", "編集者WindowsID", "資産番号", "対象項目", "変更前", "変更後", "変更理由")
End Sub

Private Sub SetupSequenceHeaders()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SEQ)

    ws.Range("A1:B1").Value = Array("採番キー", "最新連番")
End Sub

Public Function RegisterAsset(ByVal assetTypeCode As String, ByVal purchaseOrLease As String, ByVal locationCode As String, ByVal assetName As String, ByVal modelName As String, ByVal acquiredDate As Variant, ByVal amount As Variant, ByVal usefulLife As Variant, ByVal deprMethod As String, ByVal deprRate As Variant, ByVal mgmtNo As String, ByVal remarks As String, Optional ByVal statusValue As String = "使用中") As String
    Dim ws As Worksheet
    Dim nextRow As Long
    Dim assetNo As String

    ValidateRequiredFields assetTypeCode, purchaseOrLease, locationCode, assetName, modelName, acquiredDate
    ValidateLookupValues assetTypeCode, purchaseOrLease, locationCode, deprMethod
    ValidateNumericFields amount, usefulLife, deprRate

    assetNo = GenerateAssetNumber(locationCode, assetTypeCode, CDate(acquiredDate))

    Set ws = ThisWorkbook.Worksheets(SHEET_LEDGER)
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = assetNo
    ws.Cells(nextRow, 2).Value = assetTypeCode
    ws.Cells(nextRow, 3).Value = purchaseOrLease
    ws.Cells(nextRow, 4).Value = locationCode
    ws.Cells(nextRow, 5).Value = assetName
    ws.Cells(nextRow, 6).Value = modelName
    ws.Cells(nextRow, 7).Value = CDate(acquiredDate)
    ws.Cells(nextRow, 8).Value = amount
    ws.Cells(nextRow, 9).Value = usefulLife
    ws.Cells(nextRow, 10).Value = deprMethod
    ws.Cells(nextRow, 11).Value = deprRate
    ws.Cells(nextRow, 12).Value = mgmtNo
    ws.Cells(nextRow, 13).Value = remarks
    ws.Cells(nextRow, 14).Value = statusValue
    ws.Cells(nextRow, 15).Value = Now
    ws.Cells(nextRow, 16).Value = Application.UserName
    ws.Cells(nextRow, 17).Value = GetWindowsIdentity()

    RegisterAsset = assetNo
End Function
