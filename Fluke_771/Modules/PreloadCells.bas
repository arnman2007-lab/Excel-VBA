Attribute VB_Name = "PreloadCells"
' Helper function to check if a Variant contains an initialized array
Function IsArrayInitialized(v As Variant) As Boolean
    On Error GoTo ErrHandler
    If IsArray(v) Then
        ' LBound will error if array is uninitialized
        If Not IsError(LBound(v)) Then
            IsArrayInitialized = True
            Exit Function
        End If
    End If
ErrHandler:
    IsArrayInitialized = False
End Function

' Main Preload sub
Sub Preload()
    Dim j As Long
    Dim rangePair As Variant
    Dim startRow As Long
    Dim endRow As Long
    Dim rangeSplit() As String
    Dim i As Long
    Dim colLetter As String
    Dim arr As Variant
    Dim SheetName As String
    Dim k As Long
    
    ' Make sure sheet setup is done
    SetupWS
    
    ' Always initialize ranges dictionary
    ArraySetup
    
    ' Get active sheet name
    SheetName = ActiveSheet.Name
    
    ' Map the sheet to the correct ranges array using Tab variables
    Select Case SheetName
        Case Tab1
            If ranges.Exists(Tab1) Then arr = ranges(Tab1)
        Case Tab2
            If ranges.Exists(Tab2) Then arr = ranges(Tab2)
        Case Tab3
            If ranges.Exists(Tab3) Then arr = ranges(Tab3)
        Case Else
            MsgBox "No ranges defined for this sheet: " & SheetName, vbExclamation
            Exit Sub
    End Select
    
    ' Ensure arr is a proper array before looping
    If Not IsArrayInitialized(arr) Then
        MsgBox "No valid ranges for sheet: " & SheetName, vbExclamation
        Exit Sub
    End If
    
    i = 0
    
    ' ----- Information Tab setup -----
    wsInfo.Range("H13").Value = "123456-789"
    i = i + 1
    wsInfo.Range("X3").Value = make
    wsInfo.Range("Y3").Value = Model
    wsInfo.Range("W4").Value = UnitDesc
    wsInfo.Range("H14").Value = i: i = i + 1
    wsInfo.Range("H15").Value = i: i = i + 1
    wsInfo.Range("H16").Value = i: i = i + 1
    
    ' ----- Preload data into the worksheet -----
    For k = LBound(arr) To UBound(arr)
        rangeSplit = Split(arr(k), ":")
        startRow = CLng(rangeSplit(0))
        endRow = CLng(rangeSplit(1))
        
        ' Loop through each row in the range pair
        For j = startRow To endRow
            ' Loop through each column
            For c = LBound(PreloadCols) To UBound(PreloadCols)
                colLetter = PreloadCols(c)
                Worksheets(SheetName).Range(colLetter & j).Value = i
                i = i + 1
            Next c
        Next j
    Next k
    
    MsgBox i - 1
End Sub

