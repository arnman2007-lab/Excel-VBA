Attribute VB_Name = "InArray"
Function IsInArray(valueToCheck As Variant, arr As Variant) As Boolean
    Dim element As Variant
    IsInArray = False
    For Each element In arr
        If element = valueToCheck Then
            IsInArray = True
            Exit Function
        End If
    Next element
End Function
