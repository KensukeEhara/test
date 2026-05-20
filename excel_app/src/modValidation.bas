Attribute VB_Name = "modValidation"
Option Explicit

Public Sub ValidateRequiredFields(ByVal assetTypeCode As String, ByVal purchaseOrLease As String, ByVal locationCode As String, ByVal assetName As String, ByVal modelName As String, ByVal acquiredDate As Variant)
    If assetTypeCode = "" Then Err.Raise vbObjectError + 2001, , "資産種別は必須です。"
    If purchaseOrLease = "" Then Err.Raise vbObjectError + 2002, , "購入orリースは必須です。"
    If locationCode = "" Then Err.Raise vbObjectError + 2003, , "設置場所は必須です。"
    If assetName = "" Then Err.Raise vbObjectError + 2004, , "名称は必須です。"
    If modelName = "" Then Err.Raise vbObjectError + 2005, , "型式は必須です。"
    If Not IsDate(acquiredDate) Then Err.Raise vbObjectError + 2006, , "取得年月日は必須です。"
End Sub

Public Sub ValidateLookupValues(ByVal assetTypeCode As String, ByVal purchaseOrLease As String, ByVal locationCode As String, ByVal deprMethod As String)
    If Not MasterCodeExists(SHEET_TYPE_MASTER, assetTypeCode) Then
        Err.Raise vbObjectError + 2101, , "資産種別コードが不正です。"
    End If

    If Not (purchaseOrLease = "購入" Or purchaseOrLease = "リース") Then
        Err.Raise vbObjectError + 2102, , "購入orリースは購入/リースのみ設定可能です。"
    End If

    If Not MasterCodeExists(SHEET_LOCATION_MASTER, locationCode) Then
        Err.Raise vbObjectError + 2103, , "設置場所コードが不正です。"
    End If

    If deprMethod <> "" Then
        If Not DeprMethodExists(deprMethod) Then
            Err.Raise vbObjectError + 2104, , "償却方法が不正です。"
        End If
    End If
End Sub
