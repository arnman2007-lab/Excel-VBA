Attribute VB_Name = "ResetCell"
Sub ResetCells(tabName As String)
    Dim i As Long, j As Long, c As Long, k As Long
    Dim rangePair As Variant
    Dim startRow As Long, endRow As Long
    Dim rangeSplit() As String
    Dim colLetter As String
    Dim arr As Variant
    Dim totalRanges As Long, midRangeIndex As Long

    ' Setup sheet names and ranges
    SetupWS
    ArraySetup  ' Always initialize ranges dictionary

    ' Map the tabName to the correct array from the dictionary
    Select Case tabName
        Case Tab1
            If ranges.Exists(Tab1) Then arr = ranges(Tab1)
        Case Tab2
            If ranges.Exists(Tab2) Then arr = ranges(Tab2)
        Case Tab3
            If ranges.Exists(Tab3) Then arr = ranges(Tab3)
        Case Tab4
            If ranges.Exists(Tab4) Then arr = ranges(Tab4)
        Case Else
            MsgBox "No ranges defined for this sheet: " & tabName, vbExclamation
            Exit Sub
    End Select

    ' Safety check
    If Not IsArrayInitialized(arr) Then
        MsgBox "No valid ranges for sheet: " & tabName, vbExclamation
        Exit Sub
    End If

    ' Reset WorkOrderSheet info
    WorkOrderSheet.Range("H13").Value = ""
    WorkOrderSheet.Range("X3").Value = make
    WorkOrderSheet.Range("Y3").Value = Model
    WorkOrderSheet.Range("W4").Value = UnitDesc
    WorkOrderSheet.Range("H14").Value = ""
    WorkOrderSheet.Range("H15").Value = ""
    WorkOrderSheet.Range("H16").Value = "N/A"

    ' Determine midpoint of the array to split blank/N/A if needed
    totalRanges = UBound(arr) - LBound(arr) + 1
    midRangeIndex = LBound(arr) + (totalRanges \ 2) - 1

    ' Loop through each range in arr
    For k = LBound(arr) To UBound(arr)
        rangePair = arr(k)
        rangeSplit = Split(rangePair, ":")
        startRow = CLng(rangeSplit(0))
        endRow = CLng(rangeSplit(1))

        ' Loop through each row in the range
        For j = startRow To endRow
            Select Case UBound(PreloadCols) - LBound(PreloadCols)
                Case 0
                    ' 1 column only
                    colLetter = PreloadCols(0)
                    If k <= midRangeIndex Then
                        Worksheets(tabName).Range(colLetter & j).Value = ""
                    Else
                        Worksheets(tabName).Range(colLetter & j).Value = "N/A"
                    End If
                Case 1
                    ' 2 columns: first blank, second N/A
                    Worksheets(tabName).Range(PreloadCols(0) & j).Value = ""
                    Worksheets(tabName).Range(PreloadCols(1) & j).Value = "N/A"
                Case Else
                    ' More than 2 columns: first blank, rest N/A
                    For c = LBound(PreloadCols) To UBound(PreloadCols)
                        colLetter = PreloadCols(c)
                        If c = LBound(PreloadCols) Then
                            Worksheets(tabName).Range(colLetter & j).Value = ""
                        Else
                            Worksheets(tabName).Range(colLetter & j).Value = "N/A"
                        End If
                    Next c
            End Select
        Next j
    Next k
End Sub

