Attribute VB_Name = "PreloadCells"
Sub Preload(tabName As String)
    Dim j As Long
    Dim rangePair As Variant
    Dim startRow As Long
    Dim endRow As Long
    Dim rangeSplit() As String
    Dim afCell As Range, alCell As Range

    SetupWS

    If Not IsArray(ranges) Or IsEmpty(ranges) Then
        ArraySetup
    End If

    If tabName = Tab1 Then
        i = 0
        
        ' Information Tab
        WorkOrderSheet.Range("H13").Value = "5454555-001"
        i = i + 1
        
        WorkOrderSheet.Range("X3").Value = Make
        WorkOrderSheet.Range("Y3").Value = Model
        WorkOrderSheet.Range("W4").Value = UnitDesc
        
        WorkOrderSheet.Range("H14").Value = i
        i = i + 1
        WorkOrderSheet.Range("H15").Value = i
        i = i + 1
        WorkOrderSheet.Range("H16").Value = i
        i = i + 1
        
        ' Main loop through range pairs
        For Each rangePair In ranges
            rangeSplit = Split(rangePair, ":")
            startRow = CLng(rangeSplit(0))
            endRow = CLng(rangeSplit(1))
            
            For j = startRow To endRow
                ' Handle AF (As Found)
                Set afCell = dataSheet.Range(ColLetterAF & j)
                If Not afCell.MergeCells Or afCell.Address = afCell.MergeArea.Cells(1, 1).Address Then
                    afCell.Value = i
                    i = i + 1
                End If
                
                ' Handle AL (As Left)
                Set alCell = dataSheet.Range(ColLetterAL & j)
                If Not alCell.MergeCells Or alCell.Address = alCell.MergeArea.Cells(1, 1).Address Then
                    alCell.Value = i
                    i = i + 1
                End If
            Next j
        Next rangePair
        
        MsgBox i - 1
    End If
End Sub

