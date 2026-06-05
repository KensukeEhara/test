Attribute VB_Name = "modValidation"
Option Explicit

Public Sub ValidateRequiredFields(ByVal assetTypeCode As String, ByVal purchaseOrLease As String, ByVal locationCode As String, ByVal assetName As String, ByVal modelName As String, ByVal acquiredDate As Variant)
    If Trim$(assetTypeCode) = "" Then Err.Raise vbObjectError + 2001, , "資産種別は必須です。"
    If Trim$(purchaseOrLease) = "" Then Err.Raise vbObjectError + 2002, , "購入orリースは必須です。"
    If Trim$(locationCode) = "" Then Err.Raise vbObjectError + 2003, , "設置場所は必須です。"
    If Trim$(assetName) = "" Then Err.Raise vbObjectError + 2004, , "名称は必須です。"
    If Trim$(modelName) = "" Then Err.Raise vbObjectError + 2005, , "型式は必須です。"
    ValidateRequiredDate acquiredDate
End Sub

Public Sub ValidateNumericFields(ByVal amount As Variant, ByVal usefulLife As Variant, ByVal deprRate As Variant)
    If Not IsBlankValue(amount) Then
        If Not IsNumeric(amount) Then Err.Raise vbObjectError + 2201, , "取得額は数値で入力してください。"
        If CDbl(amount) < 0 Then Err.Raise vbObjectError + 2202, , "取得額は0以上で入力してください。"
    End If

    If Not IsBlankValue(usefulLife) Then
        If Not IsNumeric(usefulLife) Then Err.Raise vbObjectError + 2203, , "耐用年数は数値で入力してください。"
        If CDbl(usefulLife) <> Fix(CDbl(usefulLife)) Or CDbl(usefulLife) < 1 Then Err.Raise vbObjectError + 2204, , "耐用年数は1以上の整数で入力してください。"
    End If

    If Not IsBlankValue(deprRate) Then
        If Not IsNumeric(deprRate) Then Err.Raise vbObjectError + 2205, , "償却率は数値で入力してください。"
        If CDbl(deprRate) < 0 Or CDbl(deprRate) > 1 Then Err.Raise vbObjectError + 2206, , "償却率は0%から100%の範囲で入力してください。"
    End If
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

Private Sub ValidateRequiredDate(ByVal acquiredDate As Variant)
    If IsBlankValue(acquiredDate) Then Err.Raise vbObjectError + 2006, , "取得年月日は必須です。"
    If Not IsDate(acquiredDate) Then Err.Raise vbObjectError + 2006, , "取得年月日は必須です。"
    If CDate(acquiredDate) = 0 Then Err.Raise vbObjectError + 2006, , "取得年月日は必須です。"
End Sub

Private Function IsBlankValue(ByVal value As Variant) As Boolean
    If IsEmpty(value) Then
        IsBlankValue = True
    ElseIf IsNull(value) Then
        IsBlankValue = True
    ElseIf VarType(value) = vbString Then
        IsBlankValue = (Trim$(CStr(value)) = "")
    End If
End Function
